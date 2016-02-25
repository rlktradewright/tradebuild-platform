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
Event Error(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
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

Private mSummary                    As OrderSummary

Private mContractSpecifier          As IContractSpecifier

Private mOrderContext               As OrderContext

Private mGroupName                  As String

Private mIsDirty                    As Boolean

Private mIsSimulated                As Boolean

Private mNeedsRecovery              As Boolean

Private mBrokerData                 As Variant

Private mIsInitialised              As Boolean

Private mRecoverableObjectListeners As New Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSummary = New OrderSummary
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
IOrder_Action = mSummary.Action
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
IOrder_AllOrNone = mSummary.AllOrNone
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
IOrder_BrokerId = mSummary.BrokerId
End Property

Private Property Get IOrder_ContextsName() As String
IOrder_ContextsName = mSummary.ContextsName
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
IOrder_DiscretionaryAmount = mSummary.DiscretionaryAmount
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
IOrder_DisplaySize = mSummary.DisplaySize
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
Const ProcName As String = "IOrder_ETradeOnly"
On Error GoTo Err

ETradeOnly = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IOrder_ETradeOnly() As Boolean
IOrder_ETradeOnly = mSummary.ETradeOnly
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
IOrder_FirmQuoteOnly = mSummary.FirmQuoteOnly
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
IOrder_GoodAfterTime = mSummary.GoodAfterTime
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
IOrder_GoodAfterTimeTZ = mSummary.GoodAfterTimeTZ
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
IOrder_GoodTillDate = mSummary.GoodTillDate
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
IOrder_GoodTillDateTZ = mSummary.GoodTillDateTZ
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
IOrder_Hidden = mSummary.Hidden
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
IOrder_Id = mSummary.Id
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
IOrder_IgnoreRegularTradingHours = mSummary.IgnoreRegularTradingHours
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
IOrder_LimitPrice = mSummary.LimitPrice
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
IOrder_MinimumQuantity = mSummary.MinimumQuantity
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
IOrder_NbboPriceCap = mSummary.NbboPriceCap
End Property

Private Sub IOrder_NotifyError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
NotifyError pErrorCode, pErrorMsg
End Sub

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
IOrder_OrderType = mSummary.OrderType
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
IOrder_Origin = mSummary.Origin
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
IOrder_OriginatorRef = mSummary.OriginatorRef
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
IOrder_OverrideConstraints = mSummary.OverrideConstraints
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
Set IOrder_ProviderProperties = mSummary.ProviderProperties
End Property

Private Property Let IOrder_Quantity(ByVal Value As Long)
Quantity = Value
End Property

Private Property Get IOrder_Quantity() As Long
Const ProcName As String = "IOrder_Quantity"
On Error GoTo Err

IOrder_Quantity = mSummary.Quantity

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
IOrder_QuantityFilled = mSummary.QuantityFilled
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
IOrder_QuantityRemaining = mSummary.QuantityRemaining
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
IOrder_SettlingFirm = mSummary.SettlingFirm
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
IOrder_Status = mSummary.Status
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
IOrder_SweepToFill = mSummary.SweepToFill
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
IOrder_TimeInForce = mSummary.TimeInForce
End Property

Private Property Get IOrder_StopTriggerMethod() As OrderStopTriggerMethods
IOrder_StopTriggerMethod = mSummary.StopTriggerMethod
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
IOrder_TriggerPrice = mSummary.TriggerPrice
End Property

'@================================================================================
' IPositionMember Interface Members
'@================================================================================

Private Property Get IPositionMember_PositionManagerName() As String
IPositionMember_PositionManagerName = mSummary.ContextsName
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

Set mSummary = pData
mSummary.setClean

mNeedsRecovery = IsActive

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IRecoverable_ProgId() As String
IRecoverable_ProgId = ProjectName & "." & ModuleName
End Property

Private Property Get IRecoverable_RecoveryData() As Variant
Set IRecoverable_RecoveryData = mSummary
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

If Value = mSummary.Action Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttAction
setDirty
mSummary.Action = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Action() As OrderActions
Action = mSummary.Action
End Property

Public Property Let AllOrNone(ByVal Value As Boolean)
Const ProcName As String = "AllOrNone"
On Error GoTo Err

If Value = mSummary.AllOrNone Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttAllOrNone
setDirty
mSummary.AllOrNone = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get AllOrNone() As Boolean
AllOrNone = mSummary.AllOrNone
End Property

Public Property Let AveragePrice(ByVal Value As Double)
mSummary.AveragePrice = Value
End Property

Public Property Get AveragePrice() As Double
AveragePrice = mSummary.AveragePrice
End Property

Public Property Let BlockOrder(ByVal Value As Boolean)
Const ProcName As String = "BlockOrder"
On Error GoTo Err

If Value = mSummary.BlockOrder Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttBlockOrder
setDirty
mSummary.BlockOrder = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BlockOrder() As Boolean
BlockOrder = mSummary.BlockOrder
End Property

Public Property Let BrokerId(ByVal Value As String)
mSummary.BrokerId = Value
End Property

Public Property Get BrokerId() As String
BrokerId = mSummary.BrokerId
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Public Property Let DiscretionaryAmount(ByVal Value As Double)
Const ProcName As String = "DiscretionaryAmount"
On Error GoTo Err

If Value = mSummary.DiscretionaryAmount Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttDiscretionaryAmount
setDirty
mSummary.DiscretionaryAmount = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DiscretionaryAmount() As Double
DiscretionaryAmount = mSummary.DiscretionaryAmount
End Property

Public Property Let DisplaySize(ByVal Value As Long)
Const ProcName As String = "DisplaySize"
On Error GoTo Err

If Value = mSummary.DisplaySize Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttDisplaySize
setDirty
mSummary.DisplaySize = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mSummary.DisplaySize
End Property

Public Property Let ErrorCode(ByVal Value As Long)
mSummary.ErrorCode = Value
End Property

Public Property Get ErrorCode() As Long
ErrorCode = mSummary.ErrorCode
End Property

Private Property Let ErrorMessage(ByVal Value As String)
mSummary.ErrorMessage = Value
End Property

Private Property Get ErrorMessage() As String
ErrorMessage = mSummary.ErrorMessage
End Property

Public Property Let ETradeOnly(ByVal Value As Boolean)
Const ProcName As String = "ETradeOnly"
On Error GoTo Err

If Value = mSummary.ETradeOnly Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttETradeOnly
setDirty
mSummary.ETradeOnly = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ETradeOnly() As Boolean
ETradeOnly = mSummary.ETradeOnly
End Property

Public Property Let FillTime(ByVal Value As Date)
mSummary.FillTime = Value
End Property

Public Property Get FillTime() As Date
FillTime = mSummary.FillTime
End Property

Public Property Let FirmQuoteOnly(ByVal Value As Boolean)
Const ProcName As String = "FirmQuoteOnly"
On Error GoTo Err

If Value = mSummary.FirmQuoteOnly Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttFirmQuoteOnly
setDirty
mSummary.FirmQuoteOnly = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FirmQuoteOnly() As Boolean
FirmQuoteOnly = mSummary.FirmQuoteOnly
End Property

Public Property Let GoodAfterTime(ByVal Value As Date)
Const ProcName As String = "GoodAfterTime"
On Error GoTo Err

If Value = mSummary.GoodAfterTime Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttGoodAfterTime
setDirty
mSummary.GoodAfterTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GoodAfterTime() As Date
GoodAfterTime = mSummary.GoodAfterTime
End Property

Public Property Let GoodAfterTimeTZ(ByVal Value As String)
Const ProcName As String = "GoodAfterTimeTZ"
On Error GoTo Err

If Value = mSummary.GoodAfterTimeTZ Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttGoodAfterTimeTZ
setDirty
mSummary.GoodAfterTimeTZ = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GoodAfterTimeTZ() As String
GoodAfterTimeTZ = mSummary.GoodAfterTimeTZ
End Property

Public Property Let GoodTillDate(ByVal Value As Date)
Const ProcName As String = "GoodTillDate"
On Error GoTo Err

If Value = mSummary.GoodTillDate Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttGoodTillDate
setDirty
mSummary.GoodTillDate = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GoodTillDate() As Date
GoodTillDate = mSummary.GoodTillDate
End Property

Public Property Let GoodTillDateTZ(ByVal Value As String)
Const ProcName As String = "GoodTillDateTZ"
On Error GoTo Err

If Value = mSummary.GoodTillDateTZ Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttGoodTillDateTZ
setDirty
mSummary.GoodTillDateTZ = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GoodTillDateTZ() As String
GoodTillDateTZ = mSummary.GoodTillDateTZ
End Property

Public Property Get GroupName() As String
GroupName = mGroupName
End Property

Public Property Let Hidden(ByVal Value As Boolean)
Const ProcName As String = "Hidden"
On Error GoTo Err

If Value = mSummary.Hidden Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttHidden
setDirty
mSummary.Hidden = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Hidden() As Boolean
Hidden = mSummary.Hidden
End Property

Public Property Let Id(ByVal Value As String)
mSummary.Id = Value
End Property

Public Property Get Id() As String
Id = mSummary.Id
End Property

Public Property Let IgnoreRegularTradingHours(ByVal Value As Boolean)
Const ProcName As String = "IgnoreRegularTradingHours"
On Error GoTo Err

If Value = mSummary.IgnoreRegularTradingHours Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttIgnoreRTH
setDirty
mSummary.IgnoreRegularTradingHours = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IgnoreRegularTradingHours() As Boolean
IgnoreRegularTradingHours = mSummary.IgnoreRegularTradingHours
End Property

Public Property Get IsActive() As Boolean
IsActive = (mSummary.Status = OrderStatusPendingSubmit) Or _
        (mSummary.Status = OrderStatusPreSubmitted) Or _
        (mSummary.Status = OrderStatusSubmitted) Or _
        (mSummary.Status = OrderStatusCancelling)
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
mSummary.LastFillPrice = Value
End Property

Public Property Get LastFillPrice() As Double
LastFillPrice = mSummary.LastFillPrice
End Property

Public Property Let LimitPrice(ByVal Value As Double)
Const ProcName As String = "LimitPrice"
On Error GoTo Err

If Value = mSummary.LimitPrice Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttLimitPrice
setDirty
mSummary.LimitPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mSummary.LimitPrice
End Property

Public Property Let MinimumQuantity(ByVal Value As Long)
Const ProcName As String = "MinimumQuantity"
On Error GoTo Err

If Value = mSummary.MinimumQuantity Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttMinimumQuantity
setDirty
mSummary.MinimumQuantity = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MinimumQuantity() As Long
MinimumQuantity = mSummary.MinimumQuantity
End Property

Public Property Let NbboPriceCap(ByVal Value As Double)
Const ProcName As String = "NbboPriceCap"
On Error GoTo Err

If Value = mSummary.NbboPriceCap Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttNBBOPriceCap
setDirty
mSummary.NbboPriceCap = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mSummary.NbboPriceCap
End Property

Friend Property Get NeedsRecovery() As Boolean
NeedsRecovery = mNeedsRecovery
End Property

Friend Property Let Offset(ByVal Value As Long)
mSummary.Offset = Value
End Property

Friend Property Get Offset() As Long
Offset = mSummary.Offset
End Property

Public Property Get OrderContext() As OrderContext
Set OrderContext = mOrderContext
End Property

Public Property Let OrderType(ByVal Value As OrderTypes)
Const ProcName As String = "OrderType"
On Error GoTo Err

If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsOrderTypeSupported(Value) Then attributeValueNotSupported OrderAttributes.OrderAttOrderType, Value
End If
If Value = mSummary.OrderType Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttOrderType
setDirty
mSummary.OrderType = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mSummary.OrderType
End Property

Public Property Let Origin(ByVal Value As OrderOrigins)
Const ProcName As String = "Origin"
On Error GoTo Err

If Value = mSummary.Origin Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttOrigin
setDirty
mSummary.Origin = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Origin() As OrderOrigins
Origin = mSummary.Origin
End Property

Public Property Let OriginatorRef(ByVal Value As String)
Const ProcName As String = "OriginatorRef"
On Error GoTo Err

If Value = mSummary.OriginatorRef Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttOriginatorRef
setDirty
mSummary.OriginatorRef = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OriginatorRef() As String
OriginatorRef = mSummary.OriginatorRef
End Property

Public Property Let OverrideConstraints(ByVal Value As Boolean)
Const ProcName As String = "OverrideConstraints"
On Error GoTo Err

If Value = mSummary.OverrideConstraints Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttOverrideConstraints
setDirty
mSummary.OverrideConstraints = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OverrideConstraints() As Boolean
OverrideConstraints = mSummary.OverrideConstraints
End Property

Public Property Let PercentOffset(ByVal Value As Double)
Const ProcName As String = "PercentOffset"
On Error GoTo Err

If Value = mSummary.PercentOffset Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttPercentOffset
setDirty
mSummary.PercentOffset = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PercentOffset() As Double
PercentOffset = mSummary.PercentOffset
End Property

Public Property Let ProviderProperties( _
                ByVal Value As Parameters)
Set mSummary.ProviderProperties = Value
End Property

Public Property Get ProviderProperties() As Parameters
Set ProviderProperties = mSummary.ProviderProperties
End Property

Public Property Let Quantity(ByVal Value As Long)
Const ProcName As String = "Quantity"
On Error GoTo Err

If Value = mSummary.Quantity Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttQuantity
mSummary.Quantity = Value
setDirty
mSummary.QuantityRemaining = mSummary.Quantity - mSummary.QuantityFilled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Quantity() As Long
Quantity = mSummary.Quantity
End Property

Public Property Let QuantityFilled(ByVal Value As Long)
mSummary.QuantityFilled = Value
End Property

Public Property Get QuantityFilled() As Long
QuantityFilled = mSummary.QuantityFilled
End Property

Public Property Let QuantityRemaining(ByVal Value As Long)
mSummary.QuantityRemaining = Value
End Property

Public Property Get QuantityRemaining() As Long
QuantityRemaining = mSummary.QuantityRemaining
End Property

Public Property Let SettlingFirm(ByVal Value As String)
Const ProcName As String = "SettlingFirm"
On Error GoTo Err

If Value = mSummary.SettlingFirm Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttSettlingFirm
setDirty
mSummary.SettlingFirm = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SettlingFirm() As String
SettlingFirm = mSummary.SettlingFirm
End Property

Public Property Let Status(ByVal Value As OrderStatuses)
Const ProcName As String = "Status"
On Error GoTo Err

If Value = mSummary.Status Then Exit Property

mSummary.Status = Value
If mSummary.Status = OrderStatusFilled Or mSummary.Status = OrderStatusCancelled Then
    mSummary.ErrorCode = 0
    mSummary.ErrorMessage = ""
End If
If Not IsActive Then mNeedsRecovery = False
SaveRecoveryInfo
RaiseEvent StatusChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Status() As OrderStatuses
Status = mSummary.Status
End Property

Public Property Let StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
Const ProcName As String = "StopTriggerMethod"
On Error GoTo Err

If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsStopTriggerMethodSupported(Value) Then attributeValueNotSupported OrderAttributes.OrderAttStopTriggerMethod, Value
End If
If Value = mSummary.StopTriggerMethod Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttStopTriggerMethod
setDirty
mSummary.StopTriggerMethod = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StopTriggerMethod() As OrderStopTriggerMethods
StopTriggerMethod = mSummary.StopTriggerMethod
End Property

Public Property Let SweepToFill(ByVal Value As Boolean)
Const ProcName As String = "SweepToFill"
On Error GoTo Err

If Value = mSummary.SweepToFill Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttSweepToFill
setDirty
mSummary.SweepToFill = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSummary.SweepToFill
End Property

Public Property Let TimeInForce(ByVal Value As OrderTIFs)
Const ProcName As String = "TimeInForce"
On Error GoTo Err

AssertArgument gIsValidTIF(Value), "Value must be a member of the OrderTifs enum"

If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsOrderTifSupported(Value) Then attributeValueNotSupported OrderAttributes.OrderAttTimeInForce, Value
End If
If Value = mSummary.TimeInForce Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttTimeInForce
setDirty
mSummary.TimeInForce = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimeInForce() As OrderTIFs
TimeInForce = mSummary.TimeInForce
End Property

Public Property Let TriggerPrice(ByVal Value As Double)
Const ProcName As String = "TriggerPrice"
On Error GoTo Err

If Value = mSummary.TriggerPrice Then Exit Property
checkAttributeModificationPermitted OrderAttributes.OrderAttTriggerPrice
setDirty
mSummary.TriggerPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TriggerPrice() As Double
TriggerPrice = mSummary.TriggerPrice
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
            FormatTimestamp(pExec.FillTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "," & _
            pExec.TimezoneName & "," & _
            pExec.Id & "," & _
            pExec.BrokerId, _
            mIsSimulated, _
            Me
End If

mSummary.AveragePrice = (mSummary.QuantityFilled * mSummary.AveragePrice + pExec.Quantity * pExec.Price) / (mSummary.QuantityFilled + pExec.Quantity)
mSummary.QuantityFilled = mSummary.QuantityFilled + pExec.Quantity
mSummary.QuantityRemaining = mSummary.QuantityRemaining - pExec.Quantity
mSummary.LastFillPrice = pExec.Price
mSummary.FillTime = pExec.FillTime
If mSummary.QuantityRemaining = 0 Then
    mSummary.Status = OrderStatusFilled
    mNeedsRecovery = False
End If
SaveRecoveryInfo
RaiseEvent Fill(pExec)
If mSummary.Status = OrderStatusFilled Then RaiseEvent StatusChanged

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
                            pContextsName = mSummary.ContextsName And _
                            pContractSpec.Equals(mContractSpecifier) And _
                            pOrderContext Is mOrderContext, "Attempt to change initialisation values"

mGroupName = pGroupName
mSummary.ContextsName = pContextsName
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

Select Case mSummary.Status
Case OrderStatusCreated
    IsAttributeModifiable = True
Case OrderStatusRejected
    IsAttributeModifiable = False
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
    Select Case mSummary.OrderType
    Case OrderTypeLimit, _
        OrderTypeLimitOnClose, _
        OrderTypeStopLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeLimitOnOpen
    Case Else
        IsAttributeModifiable = False
    End Select
Case OrderAttributes.OrderAttTriggerPrice
    Select Case mSummary.OrderType
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

If mSummary.Status <> OrderStatuses.OrderStatusCancelled And _
    mSummary.Status <> OrderStatuses.OrderStatusCancelling And _
    mSummary.Status <> OrderStatuses.OrderStatusFilled _
Then IsModifiable = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub NotifyError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
ErrorCode = pErrorCode
ErrorMessage = pErrorMsg
SaveRecoveryInfo
RaiseEvent Error(pErrorCode, pErrorMsg)
End Sub

''
' Marks this order as clean, ie all changes have been either
' actioned or cancelled.
'@/
Public Sub setClean()
Const ProcName As String = "SetClean"
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
    If .Action <> mSummary.Action Then mSummary.Action = .Action: changed = True
    If .AllOrNone <> mSummary.AllOrNone Then mSummary.AllOrNone = .AllOrNone: changed = True
    If .BlockOrder <> mSummary.BlockOrder Then mSummary.BlockOrder = .BlockOrder: changed = True
    If .DiscretionaryAmount <> mSummary.DiscretionaryAmount Then mSummary.DiscretionaryAmount = .DiscretionaryAmount: changed = True
    If .DisplaySize <> mSummary.DisplaySize Then mSummary.DisplaySize = .DisplaySize: changed = True
    If .ETradeOnly <> mSummary.ETradeOnly Then mSummary.ETradeOnly = .ETradeOnly: changed = True
    If .FirmQuoteOnly <> mSummary.FirmQuoteOnly Then mSummary.FirmQuoteOnly = .FirmQuoteOnly: changed = True
    If .GoodAfterTime <> mSummary.GoodAfterTime Then mSummary.GoodAfterTime = .GoodAfterTime: changed = True
    If .GoodAfterTimeTZ <> mSummary.GoodAfterTimeTZ Then mSummary.GoodAfterTimeTZ = .GoodAfterTimeTZ: changed = True
    If .GoodTillDate <> mSummary.GoodTillDate Then mSummary.GoodTillDate = .GoodTillDate: changed = True
    If .GoodTillDateTZ <> mSummary.GoodTillDateTZ Then mSummary.GoodTillDateTZ = .GoodTillDateTZ: changed = True
    If .Hidden <> mSummary.Hidden Then mSummary.Hidden = .Hidden: changed = True
    If .IgnoreRegularTradingHours <> mSummary.IgnoreRegularTradingHours Then mSummary.IgnoreRegularTradingHours = .IgnoreRegularTradingHours: changed = True
    If .LimitPrice <> mSummary.LimitPrice Then mSummary.LimitPrice = .LimitPrice: changed = True
    If .MinimumQuantity <> mSummary.MinimumQuantity Then mSummary.MinimumQuantity = .MinimumQuantity: changed = True
    If .NbboPriceCap <> mSummary.NbboPriceCap Then mSummary.NbboPriceCap = .NbboPriceCap: changed = True
    If .OrderType <> mSummary.OrderType Then mSummary.OrderType = .OrderType: changed = True
    If .Origin <> mSummary.Origin Then mSummary.Origin = .Origin: changed = True
    If .OriginatorRef <> mSummary.OriginatorRef Then mSummary.OriginatorRef = .OriginatorRef: changed = True
    If .OverrideConstraints <> mSummary.OverrideConstraints Then mSummary.OverrideConstraints = .OverrideConstraints: changed = True
    If Not propertiesEqual(mSummary.ProviderProperties, .ProviderProperties) Then Set mSummary.ProviderProperties = .ProviderProperties: changed = True
    If .Quantity <> mSummary.Quantity Then mSummary.Quantity = .Quantity: changed = True
    If .SettlingFirm <> mSummary.SettlingFirm Then mSummary.SettlingFirm = .SettlingFirm: changed = True
    If .StopTriggerMethod <> mSummary.StopTriggerMethod Then mSummary.StopTriggerMethod = .StopTriggerMethod: changed = True
    If .SweepToFill <> mSummary.SweepToFill Then mSummary.SweepToFill = .SweepToFill: changed = True
    If .TimeInForce <> mSummary.TimeInForce Then mSummary.TimeInForce = .TimeInForce: changed = True
    If .TriggerPrice <> mSummary.TriggerPrice Then mSummary.TriggerPrice = .TriggerPrice: changed = True
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

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mRecoverableObjectListeners.Valid Then
    sInit = True
    sSomeListeners = mRecoverableObjectListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IRecoverableObjectListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyRecoveryDataChanged ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
If mSummary.Status = OrderStatuses.OrderStatusCreated Then Exit Sub
If Not mSummary.IsDirty Then Exit Sub

fireRecoveryDataChanged

mSummary.setClean

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

If mIsDirty Then Exit Sub

mIsDirty = True
RaiseEvent Dirty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub





