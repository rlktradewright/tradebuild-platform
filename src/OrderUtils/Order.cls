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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Implements IPositionMember
Implements IRecoverable

'@================================================================================
' Events
'@================================================================================

Event Clean()
Event Dirty()
Event Fill(ByVal pExec As IExecutionReport)
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

Private mContractSpecifier          As IContractSpecifier
Private mOrderContext               As OrderContext

Private mGroupName                  As String
Private mContextsName               As String

Private mIsDirty                    As Boolean

Private mIsSimulated                As Boolean

Private mStatus                     As OrderStatuses
Private mIdentifier                 As String

Private mFillTime                   As Date

Private mBrokerId                   As String
    
Private mAction                     As OrderActions
Private mQuantity                   As Long
Private mQuantityFilled             As Long
Private mQuantityRemaining          As Long
Private mAveragePrice               As Double
Private mLastFillPrice              As Double
Private mOrderType                  As OrderTypes
Private mLimitPrice                 As Double
Private mTriggerPrice               As Double
Private mOffset                     As Long

Private mTimeInForce                As OrderTIFs
Private mOrigin                     As OrderOrigins
Private mOriginatorRef              As String
Private mProviderProperties         As Parameters
Private mBlockOrder                 As Boolean
Private mSweepToFill                As Boolean
Private mDisplaySize                As Long
Private mStopTriggerMethod          As OrderStopTriggerMethods
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
    
Private mNeedsRecovery              As Boolean

Private mBrokerData                 As Variant

Private mIsInitialised              As Boolean

Private mRecoverableObjectListeners As New Listeners

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
' IOrder Interface Members
'@================================================================================

Private Property Let IOrder_Action(ByVal Value As OrderActions)
Const ProcName As String = "IOrder_Action"
On Error GoTo Err

Action = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_Action() As OrderActions
IOrder_Action = mAction
End Property

Private Property Let IOrder_AllOrNone(ByVal Value As Boolean)
Const ProcName As String = "IOrder_AllOrNone"
On Error GoTo Err

AllOrNone = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_AllOrNone() As Boolean
IOrder_AllOrNone = mAllOrNone
End Property

Private Property Let IOrder_AveragePrice(ByVal Value As Double)
Const ProcName As String = "IOrder_AveragePrice"
On Error GoTo Err

AveragePrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_AveragePrice() As Double
IOrder_AveragePrice = AveragePrice
End Property

Private Property Let IOrder_BlockOrder(ByVal Value As Boolean)
Const ProcName As String = "IOrder_BlockOrder"
On Error GoTo Err

BlockOrder = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_BlockOrder() As Boolean
IOrder_BlockOrder = BlockOrder
End Property

Private Property Let IOrder_BrokerId(ByVal Value As String)
checkInitialised

BrokerId = Value
End Property

Private Property Get IOrder_BrokerId() As String
IOrder_BrokerId = mBrokerId
End Property

Private Property Get IOrder_ContextsName() As String
IOrder_ContextsName = mContextsName
End Property

Private Property Get IOrder_ContractSpecifier() As IContractSpecifier
Set IOrder_ContractSpecifier = mContractSpecifier
End Property

Private Property Let IOrder_DiscretionaryAmount(ByVal Value As Double)
Const ProcName As String = "IOrder_DiscretionaryAmount"
On Error GoTo Err

DiscretionaryAmount = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_DiscretionaryAmount() As Double
IOrder_DiscretionaryAmount = mDiscretionaryAmount
End Property

Private Property Let IOrder_DisplaySize(ByVal Value As Long)
Const ProcName As String = "IOrder_DisplaySize"
On Error GoTo Err

DisplaySize = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_DisplaySize() As Long
IOrder_DisplaySize = mDisplaySize
End Property

Private Property Let IOrder_ETradeOnly(ByVal Value As Boolean)
Const ProcName As String = "IOrder_ETradeOnly"
On Error GoTo Err

ETradeOnly = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_ETradeOnly() As Boolean
IOrder_ETradeOnly = mETradeOnly
End Property

Private Property Let IOrder_FillTime(ByVal Value As Date)
Const ProcName As String = "IOrder_FillTime"
On Error GoTo Err

FillTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_FillTime() As Date
IOrder_FillTime = FillTime
End Property

Private Property Let IOrder_FirmQuoteOnly(ByVal Value As Boolean)
Const ProcName As String = "IOrder_FirmQuoteOnly"
On Error GoTo Err

FirmQuoteOnly = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_FirmQuoteOnly() As Boolean
IOrder_FirmQuoteOnly = mFirmQuoteOnly
End Property

Private Property Let IOrder_GoodAfterTime(ByVal Value As Date)
Const ProcName As String = "IOrder_GoodAfterTime"
On Error GoTo Err

GoodAfterTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_GoodAfterTime() As Date
IOrder_GoodAfterTime = mGoodAfterTime
End Property

Private Property Let IOrder_GoodAfterTimeTZ(ByVal Value As String)
Const ProcName As String = "IOrder_GoodAfterTimeTZ"
On Error GoTo Err

GoodAfterTimeTZ = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_GoodAfterTimeTZ() As String
IOrder_GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Private Property Let IOrder_GoodTillDate(ByVal Value As Date)
Const ProcName As String = "IOrder_GoodTillDate"
On Error GoTo Err

GoodTillDate = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_GoodTillDate() As Date
IOrder_GoodTillDate = mGoodTillDate
End Property

Private Property Let IOrder_GoodTillDateTZ(ByVal Value As String)
Const ProcName As String = "IOrder_GoodTillDateTZ"
On Error GoTo Err

GoodTillDateTZ = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_GoodTillDateTZ() As String
IOrder_GoodTillDateTZ = mGoodTillDateTZ
End Property

Private Property Get IOrder_GroupName() As String
Const ProcName As String = "IOrder_GroupName"
On Error GoTo Err

IOrder_GroupName = GroupName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Let IOrder_Hidden(ByVal Value As Boolean)
Const ProcName As String = "IOrder_Hidden"
On Error GoTo Err

Hidden = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_Hidden() As Boolean
IOrder_Hidden = mHidden
End Property

Private Property Let IOrder_Id(ByVal Value As String)
Const ProcName As String = "IOrder_Id"
On Error GoTo Err

Id = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_Id() As String
IOrder_Id = mIdentifier
End Property

Private Property Let IOrder_IgnoreRegularTradingHours(ByVal Value As Boolean)
Const ProcName As String = "IOrder_IgnoreRegularTradingHours"
On Error GoTo Err

IgnoreRegularTradingHours = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_IgnoreRegularTradingHours() As Boolean
IOrder_IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Private Sub IOrder_Initialise(ByVal pGroupName As String, ByVal pPositionManagerName As String, ByVal pContractSpec As IContractSpecifier, ByVal pOrderContext As OrderContext)
Const ProcName As String = "IOrder_Initialise"
On Error GoTo Err

Initialise pGroupName, pPositionManagerName, pContractSpec, pOrderContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IOrder_IsActive() As Boolean
IOrder_IsActive = IsActive
End Property

Private Function IOrder_IsAttributeModifiable(ByVal pAttribute As OrderAttributes) As Boolean
Const ProcName As String = "IOrder_IsAttributeModifiable"
On Error GoTo Err

IOrder_IsAttributeModifiable = IsAttributeModifiable(pAttribute)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IOrder_IsDirty() As Boolean
IOrder_IsDirty = IsDirty
End Property

Private Function IOrder_IsModifiable() As Boolean
IOrder_IsModifiable = IsModifiable
End Function

Private Property Let IOrder_IsSimulated(ByVal Value As Boolean)
Const ProcName As String = "IOrder_IsSimulated"
On Error GoTo Err

IsSimulated = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_IsSimulated() As Boolean
IOrder_IsSimulated = mIsSimulated
End Property

Private Property Let IOrder_LastFillPrice(ByVal Value As Double)
Const ProcName As String = "IOrder_LastFillPrice"
On Error GoTo Err

LastFillPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_LastFillPrice() As Double
IOrder_LastFillPrice = LastFillPrice
End Property

Private Property Let IOrder_LimitPrice(ByVal Value As Double)
Const ProcName As String = "IOrder_LimitPrice"
On Error GoTo Err

LimitPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_LimitPrice() As Double
IOrder_LimitPrice = mLimitPrice
End Property

Private Property Let IOrder_MinimumQuantity(ByVal Value As Long)
Const ProcName As String = "IOrder_MinimumQuantity"
On Error GoTo Err

MinimumQuantity = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_MinimumQuantity() As Long
IOrder_MinimumQuantity = mMinimumQuantity
End Property

Private Property Let IOrder_NbboPriceCap(ByVal Value As Double)
Const ProcName As String = "IOrder_NbboPriceCap"
On Error GoTo Err

NbboPriceCap = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_NbboPriceCap() As Double
IOrder_NbboPriceCap = mNBBOPriceCap
End Property

Private Sub IOrder_NotifyFill(ByVal pExec As IExecutionReport)
Const ProcName As String = "IOrder_NotifyFill"
On Error GoTo Err

NotifyFill pExec

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IOrder_Offset(ByVal Value As Long)
Const ProcName As String = "IOrder_Offset"
On Error GoTo Err

Offset = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_Offset() As Long
IOrder_Offset = Offset
End Property

Private Property Get IOrder_OrderContext() As OrderContext
Set IOrder_OrderContext = OrderContext
End Property

Private Property Let IOrder_OrderType(ByVal Value As OrderTypes)
Const ProcName As String = "IOrder_OrderType"
On Error GoTo Err

OrderType = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_OrderType() As OrderTypes
IOrder_OrderType = mOrderType
End Property

Private Property Let IOrder_Origin(ByVal Value As OrderOrigins)
Const ProcName As String = "IOrder_Origin"
On Error GoTo Err

Origin = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_Origin() As OrderOrigins
IOrder_Origin = mOrigin
End Property

Private Property Let IOrder_OriginatorRef(ByVal Value As String)
Const ProcName As String = "IOrder_OriginatorRef"
On Error GoTo Err

OriginatorRef = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_OriginatorRef() As String
IOrder_OriginatorRef = mOriginatorRef
End Property

Private Property Let IOrder_OverrideConstraints(ByVal Value As Boolean)
Const ProcName As String = "IOrder_OverrideConstraints"
On Error GoTo Err

OverrideConstraints = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_OverrideConstraints() As Boolean
IOrder_OverrideConstraints = mOverrideConstraints
End Property

Private Property Let IOrder_PercentOffset(ByVal Value As Double)
Const ProcName As String = "IOrder_PercentOffset"
On Error GoTo Err

PercentOffset = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_PercentOffset() As Double
IOrder_PercentOffset = PercentOffset
End Property

Private Property Let IOrder_ProviderProperties(ByVal Value As Parameters)
ProviderProperties = Value
End Property

Private Property Get IOrder_ProviderProperties() As Parameters
Set IOrder_ProviderProperties = mProviderProperties
End Property

Private Property Let IOrder_Quantity(ByVal Value As Long)
Quantity = Value
End Property

Private Property Get IOrder_Quantity() As Long
Const ProcName As String = "IOrder_Quantity"
On Error GoTo Err

IOrder_Quantity = mQuantity

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Let IOrder_QuantityFilled(ByVal Value As Long)
Const ProcName As String = "IOrder_QuantityFilled"
On Error GoTo Err

QuantityFilled = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_QuantityFilled() As Long
IOrder_QuantityFilled = mQuantityFilled
End Property

Private Property Let IOrder_QuantityRemaining(ByVal Value As Long)
Const ProcName As String = "IOrder_QuantityRemaining"
On Error GoTo Err

QuantityRemaining = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_QuantityRemaining() As Long
IOrder_QuantityRemaining = mQuantityRemaining
End Property

Private Property Let IOrder_SettlingFirm(ByVal Value As String)
Const ProcName As String = "IOrder_SettlingFirm"
On Error GoTo Err

SettlingFirm = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_SettlingFirm() As String
IOrder_SettlingFirm = mSettlingFirm
End Property

Private Property Let IOrder_Status(ByVal Value As OrderStatuses)
Const ProcName As String = "IOrder_Status"
On Error GoTo Err

Status = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_Status() As OrderStatuses
IOrder_Status = mStatus
End Property

Private Property Let IOrder_StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
Const ProcName As String = "IOrder_StopTriggerMethod"
On Error GoTo Err

StopTriggerMethod = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Let IOrder_SweepToFill(ByVal Value As Boolean)
Const ProcName As String = "IOrder_SweepToFill"
On Error GoTo Err

SweepToFill = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_SweepToFill() As Boolean
IOrder_SweepToFill = mSweepToFill
End Property

Private Sub IOrder_SyncToOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "IOrder_SyncToOrderReport"
On Error GoTo Err

SyncToOrderReport pOrderReport

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IOrder_TimeInForce(ByVal Value As OrderTIFs)
Const ProcName As String = "IOrder_TimeInForce"
On Error GoTo Err

TimeInForce = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_TimeInForce() As OrderTIFs
IOrder_TimeInForce = mTimeInForce
End Property

Private Property Get IOrder_StopTriggerMethod() As OrderStopTriggerMethods
IOrder_StopTriggerMethod = mStopTriggerMethod
End Property

Private Property Let IOrder_TriggerPrice(ByVal Value As Double)
Const ProcName As String = "IOrder_TriggerPrice"
On Error GoTo Err

TriggerPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_TriggerPrice() As Double
IOrder_TriggerPrice = mTriggerPrice
End Property

'@================================================================================
' IPositionMember Interface Members
'@================================================================================

Private Property Get IPositionMember_PositionManagerName() As String
IPositionMember_PositionManagerName = mContextsName
End Property

'@================================================================================
' IRecoverable Interface Members
'@================================================================================

Private Sub IRecoverable_AddRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "IRecoverable_AddRecoverableObjectListener"
On Error GoTo Err

mRecoverableObjectListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRecoverable_FinishRecovery()

End Sub

Private Sub IRecoverable_LoadRecoveryData(ByVal pData As Variant)
Const ProcName As String = "IRecoverable_LoadRecoveryData"
On Error GoTo Err

Dim lOs As OrderSummary
Set lOs = pData

mAction = lOs.Action
mAllOrNone = lOs.AllOrNone
mAveragePrice = lOs.AveragePrice
mBlockOrder = lOs.BlockOrder
mBrokerId = lOs.BrokerId
Set mContractSpecifier = lOs.ContractSpecifier
mDiscretionaryAmount = lOs.DiscretionaryAmount
mDisplaySize = lOs.DisplaySize
mETradeOnly = lOs.ETradeOnly
mFillTime = lOs.FillTime
mFirmQuoteOnly = lOs.FirmQuoteOnly
mGoodAfterTime = lOs.GoodAfterTime
mGoodAfterTimeTZ = lOs.GoodAfterTimeTZ
mGoodTillDate = lOs.GoodTillDate
mGoodTillDateTZ = lOs.GoodTillDateTZ
mHidden = lOs.Hidden
mIdentifier = lOs.Id
mIgnoreRegularTradingHours = lOs.IgnoreRegularTradingHours
mLastFillPrice = lOs.LastFillPrice
mLimitPrice = lOs.LimitPrice
mMinimumQuantity = lOs.MinimumQuantity
mNBBOPriceCap = lOs.NbboPriceCap
mOffset = lOs.Offset
mOrderType = lOs.OrderType
mOrigin = lOs.Origin
mOriginatorRef = lOs.OriginatorRef
mOverrideConstraints = lOs.OverrideConstraints
mPercentOffset = lOs.PercentOffset
Set mProviderProperties = lOs.ProviderProperties
mQuantity = lOs.Quantity
mQuantityFilled = lOs.QuantityFilled
mQuantityRemaining = lOs.QuantityRemaining
mSettlingFirm = lOs.SettlingFirm
mStatus = lOs.Status
mStopTriggerMethod = lOs.StopTriggerMethod
mSweepToFill = lOs.SweepToFill
mTimeInForce = lOs.TimeInForce
mTriggerPrice = lOs.TriggerPrice

mNeedsRecovery = IsActive

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IRecoverable_ProgId() As String
IRecoverable_ProgId = ProjectName & "." & ModuleName
End Property

Private Property Get IRecoverable_RecoveryData() As Variant
Set IRecoverable_RecoveryData = generateRecoveryData
End Property

Private Sub IRecoverable_RemoveRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "IRecoverable_RemoveRecoverableObjectListener"
On Error GoTo Err

mRecoverableObjectListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Action(ByVal Value As OrderActions)
Const ProcName As String = "Action"
On Error GoTo Err

checkInitialised

If Value = mAction Then Exit Property
setDirty
mAction = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Action() As OrderActions
Action = mAction
End Property

Public Property Let AllOrNone(ByVal Value As Boolean)
Const ProcName As String = "AllOrNone"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttAllOrNone
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

Public Property Let AveragePrice(ByVal Value As Double)
mAveragePrice = Value
End Property

Public Property Get AveragePrice() As Double
AveragePrice = mAveragePrice
End Property

Public Property Let BlockOrder(ByVal Value As Boolean)
Const ProcName As String = "BlockOrder"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttBlockOrder
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

Public Property Let BrokerId(ByVal Value As String)
mBrokerId = Value
End Property

Public Property Get BrokerId() As String
BrokerId = mBrokerId
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Public Property Let DiscretionaryAmount(ByVal Value As Double)
Const ProcName As String = "DiscretionaryAmount"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttDiscretionaryAmount
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
Const ProcName As String = "DisplaySize"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttDisplaySize
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

Public Property Let ETradeOnly(ByVal Value As Boolean)
Const ProcName As String = "ETradeOnly"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttETradeOnly
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

Public Property Let FillTime(ByVal Value As Date)
mFillTime = Value
End Property

Public Property Get FillTime() As Date
FillTime = mFillTime
End Property

Public Property Let FirmQuoteOnly(ByVal Value As Boolean)
Const ProcName As String = "FirmQuoteOnly"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttFirmQuoteOnly
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
Const ProcName As String = "GoodAfterTime"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttGoodAfterTime
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
Const ProcName As String = "GoodAfterTimeTZ"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttGoodAfterTimeTZ
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
Const ProcName As String = "GoodTillDate"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttGoodTillDate
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
Const ProcName As String = "GoodTillDateTZ"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttGoodTillDateTZ
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

Public Property Let Hidden(ByVal Value As Boolean)
Const ProcName As String = "Hidden"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttHidden
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

Public Property Let Id(ByVal Value As String)
mIdentifier = Value
End Property

Public Property Get Id() As String
Id = mIdentifier
End Property

Public Property Let IgnoreRegularTradingHours(ByVal Value As Boolean)
Const ProcName As String = "IgnoreRegularTradingHours"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttIgnoreRTH
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

Public Property Get IsActive() As Boolean
IsActive = (mStatus = OrderStatusPendingSubmit) Or _
        (mStatus = OrderStatusPreSubmitted) Or _
        (mStatus = OrderStatusSubmitted) Or _
        (mStatus = OrderStatusCancelling)
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mIsDirty
End Property

Public Property Let IsSimulated(ByVal Value As Boolean)
mIsSimulated = Value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Public Property Let LastFillPrice(ByVal Value As Double)
mLastFillPrice = Value
End Property

Public Property Get LastFillPrice() As Double
LastFillPrice = mLastFillPrice
End Property

Public Property Let LimitPrice(ByVal Value As Double)
Const ProcName As String = "LimitPrice"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttLimitPrice
If Value = mLimitPrice Then Exit Property
setDirty
mLimitPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mLimitPrice
End Property

Public Property Let MinimumQuantity(ByVal Value As Long)
Const ProcName As String = "MinimumQuantity"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttMinimumQuantity
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

checkAttributeModificationPermitted OrderAttributes.OrderAttNBBOPriceCap
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

Friend Property Get NeedsRecovery() As Boolean
NeedsRecovery = mNeedsRecovery
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

Public Property Let OrderType(ByVal Value As OrderTypes)
Const ProcName As String = "OrderType"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttOrderType
If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsOrderTypeSupported(Value) Then attributeValueNotSupported OrderAttributes.OrderAttOrderType, Value
End If
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
Const ProcName As String = "Origin"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttOrigin
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
Const ProcName As String = "OriginatorRef"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttOriginatorRef
If Value = mOriginatorRef Then Exit Property
setDirty
mOriginatorRef = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OriginatorRef() As String
OriginatorRef = mOriginatorRef
End Property

Public Property Let OverrideConstraints(ByVal Value As Boolean)
Const ProcName As String = "OverrideConstraints"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttOverrideConstraints
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
Const ProcName As String = "PercentOffset"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttPercentOffset
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

Public Property Let ProviderProperties( _
                ByVal Value As Parameters)
Set mProviderProperties = Value
End Property

Public Property Get ProviderProperties() As Parameters
Set ProviderProperties = mProviderProperties
End Property

Public Property Let Quantity(ByVal Value As Long)
Const ProcName As String = "Quantity"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttQuantity
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
Const ProcName As String = "SettlingFirm"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttSettlingFirm
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

Public Property Let Status(ByVal Value As OrderStatuses)
Const ProcName As String = "Status"
On Error GoTo Err

If Value <> mStatus Then
    mStatus = Value
    If Not IsActive Then mNeedsRecovery = False
    SaveRecoveryInfo
    RaiseEvent StatusChanged
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Status() As OrderStatuses
Status = mStatus
End Property

Public Property Let StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
Const ProcName As String = "StopTriggerMethod"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttStopTriggerMethod
If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsStopTriggerMethodSupported(Value) Then attributeValueNotSupported OrderAttributes.OrderAttStopTriggerMethod, Value
End If
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

Public Property Let SweepToFill(ByVal Value As Boolean)
Const ProcName As String = "SweepToFill"
On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttSweepToFill
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
Const ProcName As String = "TimeInForce"
On Error GoTo Err

AssertArgument gIsValidTIF(Value), "Value must be a member of the OrderTifs enum"

checkAttributeModificationPermitted OrderAttributes.OrderAttTimeInForce
If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsOrderTifSupported(Value) Then attributeValueNotSupported OrderAttributes.OrderAttTimeInForce, Value
End If
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

Public Property Let TriggerPrice(ByVal Value As Double)
mTriggerPrice = Value
End Property

Public Property Get TriggerPrice() As Double
TriggerPrice = mTriggerPrice
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Clone() As IOrder
Const ProcName As String = "Clone"
On Error GoTo Err

checkInitialised

Dim lOrder As New Order
gSyncToOrder lOrder, Me

Set Clone = lOrder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Notifies a fill for this <code>order</code>.
'
' @param pExec  an <code>Execution</code> object containing details of the fill
'
'@/
Public Sub NotifyFill( _
                ByVal pExec As IExecutionReport)
Const ProcName As String = "NotifyFill"
On Error GoTo Err

checkInitialised

If mContractSpecifier Is Nothing Then
    ' happens when an order from a previous session is being recovered before
    ' the relevant ticker is ready
Else
    gLogOrderDetail mContractSpecifier.Exchange & "," & _
            mContractSpecifier.LocalSymbol & "," & _
            pExec.Quantity & "," & _
            pExec.Price & "," & _
            gOrderActionToString(pExec.Action) & "," & _
            pExec.FillTime & "," & _
            pExec.TimezoneName & "," & _
            pExec.Id & "," & _
            pExec.BrokerId, _
            mIsSimulated, _
            Me
End If

mAveragePrice = (mQuantityFilled * mAveragePrice + pExec.Quantity * pExec.Price) / (mQuantityFilled + pExec.Quantity)
mQuantityFilled = mQuantityFilled + pExec.Quantity
mQuantityRemaining = mQuantityRemaining - pExec.Quantity
mLastFillPrice = pExec.Price
mFillTime = pExec.FillTime
If mQuantityRemaining = 0 Then
    mStatus = OrderStatusFilled
    mNeedsRecovery = False
End If
SaveRecoveryInfo
RaiseEvent Fill(pExec)
If mStatus = OrderStatusFilled Then RaiseEvent StatusChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mContractSpecifier = Nothing
Set mOrderContext = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pGroupName As String, _
                ByVal pContextsName As String, _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pOrderContext As OrderContext)
Const ProcName As String = "Initialise"
On Error GoTo Err

If mIsInitialised Then Assert pGroupName = mGroupName And _
                            pContextsName = mContextsName And _
                            pContractSpec.Equals(mContractSpecifier) And _
                            pOrderContext Is mOrderContext, "Attempt to change initialisation values"

mGroupName = pGroupName
mContextsName = pContextsName
Set mOrderContext = pOrderContext
Set mContractSpecifier = pContractSpec
mIsInitialised = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function IsAttributeModifiable( _
                ByVal pAttribute As OrderAttributes) As Boolean
Const ProcName As String = "IsAttributeModifiable"
On Error GoTo Err

If mOrderContext Is Nothing Then
    ' we are currently recovering persisted order information so don't make this check
    IsAttributeModifiable = True
    Exit Function
End If

If Not mOrderContext.IsOrderAttributeSupported(pAttribute) Then
    IsAttributeModifiable = False
    Exit Function
End If

Select Case mStatus
Case OrderStatusCreated
    IsAttributeModifiable = True
Case OrderStatusRejected
    IsAttributeModifiable = True
Case OrderStatusPendingSubmit
    IsAttributeModifiable = pAttribute And _
                            (OrderAttributes.OrderAttAllOrNone Or _
                            OrderAttributes.OrderAttBlockOrder Or _
                            OrderAttributes.OrderAttDiscretionaryAmount Or _
                            OrderAttributes.OrderAttDisplaySize Or _
                            OrderAttributes.OrderAttETradeOnly Or _
                            OrderAttributes.OrderAttFirmQuoteOnly Or _
                            OrderAttributes.OrderAttGoodAfterTime Or _
                            OrderAttributes.OrderAttGoodAfterTimeTZ Or _
                            OrderAttributes.OrderAttGoodTillDate Or _
                            OrderAttributes.OrderAttGoodTillDateTZ Or _
                            OrderAttributes.OrderAttHidden Or _
                            OrderAttributes.OrderAttIgnoreRTH Or _
                            OrderAttributes.OrderAttLimitPrice Or _
                            OrderAttributes.OrderAttMinimumQuantity Or _
                            OrderAttributes.OrderAttOriginatorRef Or _
                            OrderAttributes.OrderAttOverrideConstraints Or _
                            OrderAttributes.OrderAttPercentOffset Or _
                            OrderAttributes.OrderAttQuantity Or _
                            OrderAttributes.OrderAttRule80A Or _
                            OrderAttributes.OrderAttSettlingFirm Or _
                            OrderAttributes.OrderAttStopTriggerMethod Or _
                            OrderAttributes.OrderAttSweepToFill Or _
                            OrderAttributes.OrderAttTimeInForce Or _
                            OrderAttributes.OrderAttTriggerPrice)
Case OrderStatusPreSubmitted
    IsAttributeModifiable = pAttribute And _
                            (OrderAttributes.OrderAttAllOrNone Or _
                            OrderAttributes.OrderAttBlockOrder Or _
                            OrderAttributes.OrderAttDiscretionaryAmount Or _
                            OrderAttributes.OrderAttDisplaySize Or _
                            OrderAttributes.OrderAttETradeOnly Or _
                            OrderAttributes.OrderAttFirmQuoteOnly Or _
                            OrderAttributes.OrderAttGoodAfterTime Or _
                            OrderAttributes.OrderAttGoodAfterTimeTZ Or _
                            OrderAttributes.OrderAttGoodTillDate Or _
                            OrderAttributes.OrderAttGoodTillDateTZ Or _
                            OrderAttributes.OrderAttHidden Or _
                            OrderAttributes.OrderAttLimitPrice Or _
                            OrderAttributes.OrderAttMinimumQuantity Or _
                            OrderAttributes.OrderAttOriginatorRef Or _
                            OrderAttributes.OrderAttQuantity Or _
                            OrderAttributes.OrderAttSweepToFill Or _
                            OrderAttributes.OrderAttTimeInForce Or _
                            OrderAttributes.OrderAttTriggerPrice)
Case OrderStatusSubmitted
    IsAttributeModifiable = pAttribute And _
                            (OrderAttributes.OrderAttDiscretionaryAmount Or _
                            OrderAttributes.OrderAttDisplaySize Or _
                            OrderAttributes.OrderAttGoodTillDate Or _
                            OrderAttributes.OrderAttGoodTillDateTZ Or _
                            OrderAttributes.OrderAttHidden Or _
                            OrderAttributes.OrderAttLimitPrice Or _
                            OrderAttributes.OrderAttMinimumQuantity Or _
                            OrderAttributes.OrderAttQuantity Or _
                            OrderAttributes.OrderAttTimeInForce Or _
                            OrderAttributes.OrderAttTriggerPrice)
Case OrderStatusFilled
    IsAttributeModifiable = False
Case OrderStatusCancelling
    IsAttributeModifiable = False
Case OrderStatusCancelled
    IsAttributeModifiable = False
End Select

Select Case pAttribute
Case OrderAttributes.OrderAttLimitPrice
    Select Case mOrderType
    Case OrderTypeLimit, _
        OrderTypeLimitOnClose, _
        OrderTypeStopLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeLimitOnOpen
    Case Else
        IsAttributeModifiable = False
    End Select
Case OrderAttributes.OrderAttTriggerPrice
    Select Case mOrderType
    Case OrderTypeStop, _
        OrderTypeStopLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeMarketIfTouched
    Case Else
        IsAttributeModifiable = False
    End Select
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsModifiable() As Boolean
Const ProcName As String = "IsModifiable"

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
Public Sub setClean()
Const ProcName As String = "setClean"
On Error GoTo Err

mIsDirty = False
RaiseEvent Clean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SyncToOrderReport( _
                ByVal pExec As IOrderReport)
Const ProcName As String = "SyncToOrderReport"
On Error GoTo Err

Dim changed As Boolean

With pExec
    If .Action <> mAction Then mAction = .Action: changed = True
    If .AllOrNone <> mAllOrNone Then mAllOrNone = .AllOrNone: changed = True
    If .BlockOrder <> mBlockOrder Then mBlockOrder = .BlockOrder: changed = True
    If .DiscretionaryAmount <> mDiscretionaryAmount Then mDiscretionaryAmount = .DiscretionaryAmount: changed = True
    If .DisplaySize <> mDisplaySize Then mDisplaySize = .DisplaySize: changed = True
    If .ETradeOnly <> mETradeOnly Then mETradeOnly = .ETradeOnly: changed = True
    If .FirmQuoteOnly <> mFirmQuoteOnly Then mFirmQuoteOnly = .FirmQuoteOnly: changed = True
    If .GoodAfterTime <> mGoodAfterTime Then mGoodAfterTime = .GoodAfterTime: changed = True
    If .GoodAfterTimeTZ <> mGoodAfterTimeTZ Then mGoodAfterTimeTZ = .GoodAfterTimeTZ: changed = True
    If .GoodTillDate <> mGoodTillDate Then mGoodTillDate = .GoodTillDate: changed = True
    If .GoodTillDateTZ <> mGoodTillDateTZ Then mGoodTillDateTZ = .GoodTillDateTZ: changed = True
    If .Hidden <> mHidden Then mHidden = .Hidden: changed = True
    If .IgnoreRegularTradingHours <> mIgnoreRegularTradingHours Then mIgnoreRegularTradingHours = .IgnoreRegularTradingHours: changed = True
    If .LimitPrice <> mLimitPrice Then mLimitPrice = .LimitPrice: changed = True
    If .MinimumQuantity <> mMinimumQuantity Then mMinimumQuantity = .MinimumQuantity: changed = True
    If .NbboPriceCap <> mNBBOPriceCap Then mNBBOPriceCap = .NbboPriceCap: changed = True
    If .OrderType <> mOrderType Then mOrderType = .OrderType: changed = True
    If .Origin <> mOrigin Then mOrigin = .Origin: changed = True
    If .OriginatorRef <> mOriginatorRef Then mOriginatorRef = .OriginatorRef: changed = True
    If .OverrideConstraints <> mOverrideConstraints Then mOverrideConstraints = .OverrideConstraints: changed = True
    If Not propertiesEqual(mProviderProperties, .ProviderProperties) Then Set mProviderProperties = .ProviderProperties: changed = True
    If .Quantity <> mQuantity Then mQuantity = .Quantity: changed = True
    If .SettlingFirm <> mSettlingFirm Then mSettlingFirm = .SettlingFirm: changed = True
    If .StopTriggerMethod <> mStopTriggerMethod Then mStopTriggerMethod = .StopTriggerMethod: changed = True
    If .SweepToFill <> mSweepToFill Then mSweepToFill = .SweepToFill: changed = True
    If .TimeInForce <> mTimeInForce Then mTimeInForce = .TimeInForce: changed = True
    If .TriggerPrice <> mTriggerPrice Then mTriggerPrice = .TriggerPrice: changed = True
End With

mNeedsRecovery = False

If changed Then
    SaveRecoveryInfo
    RaiseEvent PropertyChanged
End If
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'''
'' Notifies the specified <code>order</code> of a change of status
''
'' @param pStatus    a member of the <code>OrderStatuses</code> enumeration
''                   specifying the <code>order</code>'s new status
''
''@/
'Public Sub UpdateStatus( _
'                ByVal pStatus As OrderStatuses)
'Const ProcName As String = "UpdateStatus"
'On Error GoTo Err
'
'If pStatus <> mStatus Then
'    mStatus = pStatus
'    If Not IsActive Then mNeedsRecovery = False
'    SaveRecoveryInfo
'    RaiseEvent StatusChanged
'End If
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attributeValueNotSupported( _
                ByVal pAttribute As OrderAttributes, _
                ByVal Value As Variant)
Const ProcName As String = "attributeValueNotSupported"

On Error GoTo Err

AssertArgument False, "Value " & CStr(Value) & " for attribute " & gOrderAttributeToString(pAttribute) & " is not supported by the service provider"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub checkInitialised()
Const ProcName As String = "checkInitialised"
On Error GoTo Err

Assert mIsInitialised, "Order has not been initialised"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub checkAttributeModificationPermitted( _
                ByVal attributeId As OrderAttributes)
Const ProcName As String = "checkAttributeModificationPermitted"
On Error GoTo Err

checkInitialised

Assert IsAttributeModifiable(attributeId), "The " & gOrderAttributeToString(attributeId) & " attribute is not currently modifiable"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireRecoveryDataChanged()
Const ProcName As String = "fireRecoveryDataChanged"
On Error GoTo Err

If mRecoverableObjectListeners Is Nothing Or mRecoverableObjectListeners.Count = 0 Then Exit Sub

Dim ev As RecoveryDataChangedEventData
Set ev.Source = Me

mRecoverableObjectListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mRecoverableObjectListeners.Count
    Dim lListener As IRecoverableObjectListener
    Set lListener = mRecoverableObjectListeners.GetListener(i)
    lListener.NotifyRecoveryDataChanged ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function generateRecoveryData() As Object
Const ProcName As String = "generateRecoveryData"
On Error GoTo Err

Dim summary As OrderSummary
Set summary = New OrderSummary

summary.Action = mAction
summary.AllOrNone = mAllOrNone
summary.AveragePrice = mAveragePrice
summary.BlockOrder = False
summary.BrokerId = mBrokerId
summary.ContextsName = mContextsName
summary.ContractSpecifier = mContractSpecifier
summary.DiscretionaryAmount = mDiscretionaryAmount
summary.DisplaySize = mDisplaySize
summary.ETradeOnly = mETradeOnly
summary.FillTime = mFillTime
summary.FirmQuoteOnly = mFirmQuoteOnly
summary.GoodAfterTime = mGoodAfterTime
summary.GoodAfterTimeTZ = mGoodAfterTimeTZ
summary.GoodTillDate = mGoodTillDate
summary.GoodTillDateTZ = mGoodTillDateTZ
summary.Hidden = mHidden
summary.Id = mIdentifier
summary.IgnoreRegularTradingHours = mIgnoreRegularTradingHours
summary.LastFillPrice = mLastFillPrice
summary.LimitPrice = mLimitPrice
summary.MinimumQuantity = mMinimumQuantity
summary.NbboPriceCap = mNBBOPriceCap
summary.Offset = mOffset
summary.OrderType = mOrderType
summary.Origin = mOrigin
summary.OriginatorRef = mOriginatorRef
summary.OverrideConstraints = mOverrideConstraints
summary.PercentOffset = mPercentOffset
summary.ProviderProperties = mProviderProperties
summary.Quantity = mQuantity
summary.QuantityFilled = mQuantityFilled
summary.QuantityRemaining = mQuantityRemaining
summary.SettlingFirm = mSettlingFirm
summary.Status = mStatus
summary.StopTriggerMethod = mStopTriggerMethod
summary.SweepToFill = mSweepToFill
summary.TimeInForce = mTimeInForce
summary.TriggerPrice = mTriggerPrice

Set generateRecoveryData = summary

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

Private Sub SaveRecoveryInfo()
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

If mIsSimulated Then Exit Sub
If mStatus = OrderStatuses.OrderStatusCreated Then Exit Sub

fireRecoveryDataChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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





