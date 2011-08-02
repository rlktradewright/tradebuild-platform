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

Private mContract                   As Contract
Private mOrderContext               As OrderContext

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
Private moriginatorRef              As String
Private mProviderProperties         As Parameters
Private mBlockOrder                 As Boolean
Private mSweepToFill                As Boolean
Private mDisplaySize                As Long
Private mStopTriggerMethod          As StopTriggerMethods
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

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mStatus = OrderStatuses.OrderStatusCreated
mAction = OrderActionBuy
mOrderType = OrderTypeNone
mTimeInForce = OrderTIFDay

Set mProviderProperties = New Parameters

End Sub

'@================================================================================
' IOrder Interface Members
'@================================================================================

Private Property Get IOrder_Action() As OrderActions
IOrder_Action = mAction
End Property

Private Property Get IOrder_AllOrNone() As Boolean
IOrder_AllOrNone = mAllOrNone
End Property

Private Property Get IOrder_BlockOrder() As Boolean
IOrder_BlockOrder = mBlockOrder
End Property

Private Property Let IOrder_BrokerData(ByVal value As Variant)
Const ProcName As String = "IOrder_BrokerData"
On Error GoTo Err

gSetVariant mBrokerData, value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IOrder_BrokerData() As Variant
Const ProcName As String = "IOrder_BrokerData"
On Error GoTo Err

gSetVariant IOrder_BrokerData, mBrokerData

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Let IOrder_BrokerId(ByVal RHS As String)
BrokerId = RHS
End Property

Private Property Get IOrder_BrokerId() As String
IOrder_BrokerId = mBrokerId
End Property

Private Property Get IOrder_ContractSpec() As ContractSpecifier
Set IOrder_ContractSpec = mContract.Specifier
End Property

Private Property Get IOrder_DiscretionaryAmount() As Double
IOrder_DiscretionaryAmount = mDiscretionaryAmount
End Property

Private Property Get IOrder_DisplaySize() As Long
IOrder_DisplaySize = mDisplaySize
End Property

Private Property Get IOrder_ETradeOnly() As Boolean
IOrder_ETradeOnly = mETradeOnly
End Property

Private Property Get IOrder_FirmQuoteOnly() As Boolean
IOrder_FirmQuoteOnly = mFirmQuoteOnly
End Property

Private Property Get IOrder_GoodAfterTime() As Date
IOrder_GoodAfterTime = mGoodAfterTime
End Property

Private Property Get IOrder_GoodAfterTimeTZ() As String
IOrder_GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Private Property Get IOrder_GoodTillDate() As Date
IOrder_GoodTillDate = mGoodTillDate
End Property

Private Property Get IOrder_GoodTillDateTZ() As String
IOrder_GoodTillDateTZ = mGoodTillDateTZ
End Property

Private Property Get IOrder_Hidden() As Boolean
IOrder_Hidden = mHidden
End Property

Private Property Get IOrder_Id() As String
IOrder_Id = mIdentifier
End Property

Private Property Get IOrder_IgnoreRegularTradingHours() As Boolean
IOrder_IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Private Property Get IOrder_IsSimulated() As Boolean
IOrder_IsSimulated = mIsSimulated
End Property

Private Property Get IOrder_LimitPrice() As Double
IOrder_LimitPrice = mLimitPrice
End Property

Private Property Get IOrder_MinimumQuantity() As Long
IOrder_MinimumQuantity = mMinimumQuantity
End Property

Private Property Get IOrder_NbboPriceCap() As Double
IOrder_NbboPriceCap = mNBBOPriceCap
End Property

Private Property Get IOrder_OrderType() As OrderTypes
IOrder_OrderType = mOrderType
End Property

Private Property Get IOrder_Origin() As OrderOrigins
IOrder_Origin = mOrigin
End Property

Private Property Get IOrder_OriginatorRef() As String
IOrder_OriginatorRef = moriginatorRef
End Property

Private Property Get IOrder_OverrideConstraints() As Boolean
IOrder_OverrideConstraints = mOverrideConstraints
End Property

Private Property Get IOrder_ProviderProperties() As TWUtilities30.Parameters
Set IOrder_ProviderProperties = mProviderProperties
End Property

Private Property Get IOrder_Quantity() As Long
IOrder_Quantity = mQuantity
End Property

Private Property Get IOrder_QuantityFilled() As Long
IOrder_QuantityFilled = mQuantityFilled
End Property

Private Property Get IOrder_QuantityRemaining() As Long
IOrder_QuantityRemaining = mQuantityRemaining
End Property

Private Property Get IOrder_SettlingFirm() As String
IOrder_SettlingFirm = mSettlingFirm
End Property

Private Property Get IOrder_Status() As OrderStatuses
IOrder_Status = mStatus
End Property

Private Property Get IOrder_SweepToFill() As Boolean
IOrder_SweepToFill = mSweepToFill
End Property

Private Property Get IOrder_TimeInForce() As OrderTIFs
IOrder_TimeInForce = mTimeInForce
End Property

Private Property Get IOrder_StopTriggerMethod() As StopTriggerMethods
IOrder_StopTriggerMethod = mStopTriggerMethod
End Property

Private Property Get IOrder_TriggerPrice() As Double
IOrder_TriggerPrice = mTriggerPrice
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Action(ByVal value As OrderActions)
mAction = value
End Property

Public Property Get Action() As OrderActions
Action = mAction
End Property

Public Property Let AllOrNone(ByVal value As Boolean)
Const ProcName As String = "AllOrNone"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttAllOrNone
If value = mAllOrNone Then Exit Property
setDirty
mAllOrNone = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get AllOrNone() As Boolean
AllOrNone = mAllOrNone
End Property

Friend Property Let AveragePrice(ByVal value As Double)
mAveragePrice = value
End Property

Public Property Get AveragePrice() As Double
AveragePrice = mAveragePrice
End Property

Public Property Let BlockOrder(value As Boolean)
Const ProcName As String = "BlockOrder"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttBlockOrder
If value = mBlockOrder Then Exit Property
setDirty
mBlockOrder = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BlockOrder() As Boolean
BlockOrder = mBlockOrder
End Property

Friend Property Let BrokerId(ByVal value As String)
mBrokerId = value
End Property

Public Property Get BrokerId() As String
BrokerId = mBrokerId
End Property

Public Property Let Contract(ByVal value As Contract)
Set Contract = mContract
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Public Property Let DiscretionaryAmount(ByVal value As Double)
Const ProcName As String = "DiscretionaryAmount"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttDiscretionaryAmount
If value = mDiscretionaryAmount Then Exit Property
setDirty
mDiscretionaryAmount = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DiscretionaryAmount() As Double
DiscretionaryAmount = mDiscretionaryAmount
End Property

Public Property Let DisplaySize(ByVal value As Long)
Const ProcName As String = "DisplaySize"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttDisplaySize
If value = mDisplaySize Then Exit Property
setDirty
mDisplaySize = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mDisplaySize
End Property

Public Property Let ETradeOnly(ByVal value As Boolean)
Const ProcName As String = "ETradeOnly"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttETradeOnly
If value = mETradeOnly Then Exit Property
setDirty
mETradeOnly = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ETradeOnly() As Boolean
ETradeOnly = mETradeOnly
End Property

Friend Property Let FillTime(ByVal value As Date)
mFillTime = value
End Property

Public Property Get FillTime() As Date
FillTime = mFillTime
End Property

Public Property Let FirmQuoteOnly(ByVal value As Boolean)
Const ProcName As String = "FirmQuoteOnly"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttFirmQuoteOnly
If value = mFirmQuoteOnly Then Exit Property
setDirty
mFirmQuoteOnly = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get FirmQuoteOnly() As Boolean
FirmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let GoodAfterTime(ByVal value As Date)
Const ProcName As String = "GoodAfterTime"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttGoodAfterTime
If value = mGoodAfterTime Then Exit Property
setDirty
mGoodAfterTime = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get GoodAfterTime() As Date
GoodAfterTime = mGoodAfterTime
End Property

Public Property Let GoodAfterTimeTZ(ByVal value As String)
Const ProcName As String = "GoodAfterTimeTZ"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttGoodAfterTimeTZ
If value = mGoodAfterTimeTZ Then Exit Property
setDirty
mGoodAfterTimeTZ = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get GoodAfterTimeTZ() As String
GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let GoodTillDate(ByVal value As Date)
Const ProcName As String = "GoodTillDate"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttGoodTillDate
If value = mGoodTillDate Then Exit Property
setDirty
mGoodTillDate = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get GoodTillDate() As Date
GoodTillDate = mGoodTillDate
End Property

Public Property Let GoodTillDateTZ(ByVal value As String)
Const ProcName As String = "GoodTillDateTZ"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttGoodTillDateTZ
If value = mGoodTillDateTZ Then Exit Property
setDirty
mGoodTillDateTZ = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get GoodTillDateTZ() As String
GoodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let Hidden(value As Boolean)
Const ProcName As String = "Hidden"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttHidden
If value = mHidden Then Exit Property
setDirty
mHidden = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Friend Property Let Id(ByVal value As String)
mIdentifier = value
End Property

Public Property Get Id() As String
Id = mIdentifier
End Property

Public Property Let IgnoreRegularTradingHours(value As Boolean)
Const ProcName As String = "IgnoreRegularTradingHours"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttIgnoreRTH
If value = mIgnoreRegularTradingHours Then Exit Property
setDirty
mIgnoreRegularTradingHours = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Friend Property Let IsSimulated(ByVal value As Boolean)
mIsSimulated = value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Let LastFillPrice(ByVal value As Double)
mLastFillPrice = value
End Property

Public Property Get LastFillPrice() As Double
LastFillPrice = mLastFillPrice
End Property

Friend Property Let LimitPrice(ByVal value As Double)
mLimitPrice = value
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mLimitPrice
End Property

Public Property Let MinimumQuantity(ByVal value As Long)
Const ProcName As String = "MinimumQuantity"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttMinimumQuantity
If value = mMinimumQuantity Then Exit Property
setDirty
mMinimumQuantity = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MinimumQuantity() As Long
MinimumQuantity = mMinimumQuantity
End Property

Public Property Let NbboPriceCap(ByVal value As Double)
Const ProcName As String = "NbboPriceCap"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttNBBOPriceCap
If value = mNBBOPriceCap Then Exit Property
setDirty
mNBBOPriceCap = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
End Property

Friend Property Let NeedsRecovery(ByVal value As Boolean)
mNeedsRecovery = value
End Property

Friend Property Get NeedsRecovery() As Boolean
NeedsRecovery = mNeedsRecovery
End Property

Friend Property Let Offset(ByVal value As Long)
mOffset = value
End Property

Friend Property Get Offset() As Long
Offset = mOffset
End Property

Friend Property Let OrderContext(ByVal value As OrderContext)
Set mOrderContext = value
End Property

Public Property Get OrderContext() As OrderContext
Set OrderContext = mOrderContext
End Property

Friend Property Let OrderType(ByVal value As OrderTypes)
Const ProcName As String = "OrderType"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttOrderType
If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsOrderTypeValueSupported(value) Then attributeValueNotSupported OrderAttributes.OrderAttOrderType, value
End If
If value = mOrderType Then Exit Property
setDirty
mOrderType = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mOrderType
End Property

Public Property Let Origin(ByVal value As OrderOrigins)
Const ProcName As String = "Origin"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttOrigin
If value = mOrigin Then Exit Property
setDirty
mOrigin = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Origin() As OrderOrigins
Origin = mOrigin
End Property

Public Property Let OriginatorRef(ByVal value As String)
Const ProcName As String = "OriginatorRef"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttOriginatorRef
If value = moriginatorRef Then Exit Property
setDirty
moriginatorRef = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get OriginatorRef() As String
OriginatorRef = moriginatorRef
End Property

Public Property Let OverrideConstraints(ByVal value As Boolean)
Const ProcName As String = "OverrideConstraints"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttOverrideConstraints
If value = mOverrideConstraints Then Exit Property
setDirty
mOverrideConstraints = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get OverrideConstraints() As Boolean
OverrideConstraints = mOverrideConstraints
End Property

Public Property Let PercentOffset(ByVal value As Double)
Const ProcName As String = "PercentOffset"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttPercentOffset
If value = mPercentOffset Then Exit Property
setDirty
mPercentOffset = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PercentOffset() As Double
PercentOffset = mPercentOffset
End Property

Friend Property Let ProviderProperties( _
                ByVal value As TWUtilities30.Parameters)
Set mProviderProperties = value
End Property

Public Property Get ProviderProperties() As TWUtilities30.Parameters
Set ProviderProperties = mProviderProperties
End Property

Public Property Let Quantity(ByVal value As Long)
Const ProcName As String = "Quantity"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttQuantity
If value = mQuantity Then Exit Property
mQuantity = value
setDirty
mQuantityRemaining = mQuantity - mQuantityFilled

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Friend Property Let QuantityFilled(ByVal value As Long)
mQuantityFilled = value
End Property

Public Property Get QuantityFilled() As Long
QuantityFilled = mQuantityFilled
End Property

Friend Property Let QuantityRemaining(ByVal value As Long)
mQuantityRemaining = value
End Property

Public Property Get QuantityRemaining() As Long
QuantityRemaining = mQuantityRemaining
End Property

Public Property Let SettlingFirm(value As String)
Const ProcName As String = "SettlingFirm"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttSettlingFirm
If value = mSettlingFirm Then Exit Property
setDirty
mSettlingFirm = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SettlingFirm() As String
SettlingFirm = mSettlingFirm
End Property

Friend Property Let Status(value As OrderStatuses)
mStatus = value
End Property

Public Property Get Status() As OrderStatuses
Status = mStatus
End Property

Public Property Let StopTriggerMethod(ByVal value As StopTriggerMethods)
Const ProcName As String = "StopTriggerMethod"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttStopTriggerMethod
If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsStopTriggerMethodValueSupported(value) Then attributeValueNotSupported OrderAttributes.OrderAttStopTriggerMethod, value
End If
If value = mStopTriggerMethod Then Exit Property
setDirty
mStopTriggerMethod = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get StopTriggerMethod() As StopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(value As Boolean)
Const ProcName As String = "SweepToFill"

On Error GoTo Err

checkAttributeModificationPermitted OrderAttributes.OrderAttSweepToFill
If value = mSweepToFill Then Exit Property
setDirty
mSweepToFill = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Public Property Let TimeInForce(ByVal value As OrderTIFs)
Const ProcName As String = "TimeInForce"

On Error GoTo Err

If Not gIsValidTIF(value) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must be a member of the OrderTifs enum"

checkAttributeModificationPermitted OrderAttributes.OrderAttTimeInForce
If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsOrderTifValueSupported(value) Then attributeValueNotSupported OrderAttributes.OrderAttTimeInForce, value
End If
If value = mTimeInForce Then Exit Property
setDirty
mTimeInForce = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TimeInForce() As OrderTIFs
TimeInForce = mTimeInForce
End Property

Friend Property Let TriggerPrice(ByVal value As Double)
mTriggerPrice = value
End Property

Public Property Get TriggerPrice() As Double
TriggerPrice = mTriggerPrice
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Clone() As Order
Const ProcName As String = "Clone"

On Error GoTo Err

Set Clone = New Order
syncToMe Clone

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

''
' Notifies a fill for this <code>order</code>.
'
' @param exec  an <code>Execution</code> object containing details of the fill
'
'@/
Friend Sub NotifyFill( _
                ByVal exec As IExecutionReport, _
                ByRef persistenceKey As String)

Const ProcName As String = "NotifyFill"

On Error GoTo Err

If mContract Is Nothing Then
    ' happens when an order from a previous session is being recovered before
    ' the relevant ticker is ready
Else
    gLogOrderDetail mContract.Specifier.Exchange & "," & _
            mContract.Specifier.Symbol & "," & _
            mContract.Specifier.Expiry & "," & _
            exec.Quantity & "," & _
            exec.price & "," & _
            OrderActionToString(exec.Action) & "," & _
            exec.FillTime & "," & _
            exec.TimezoneName & "," & _
            exec.Id & "," & _
            exec.BrokerId, _
            exec.IsSimulated, _
            Me
End If

mAveragePrice = (mQuantityFilled * mAveragePrice + exec.Quantity * exec.price) / (mQuantityFilled + exec.Quantity)
mQuantityFilled = mQuantityFilled + exec.Quantity
mQuantityRemaining = mQuantityRemaining - exec.Quantity
mLastFillPrice = exec.price
mFillTime = exec.FillTime
If mQuantityRemaining = 0 Then
    mStatus = OrderStatusFilled
    NeedsRecovery = False
End If
SaveRecoveryInfo persistenceKey
RaiseEvent Fill(exec)
If mStatus = OrderStatusFilled Then RaiseEvent StatusChanged

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"

On Error GoTo Err

Set mContract = Nothing
Set mOrderContext = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function IsAttributeModifiable( _
                ByVal attributeId As OrderAttributes) As Boolean

Const ProcName As String = "IsAttributeModifiable"

On Error GoTo Err

If mOrderContext Is Nothing Then
    ' we are currently recovering persisted order information so don't make this check
    IsAttributeModifiable = True
    Exit Function
End If

If Not mOrderContext.IsAttributeSupported(attributeId) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The " & GOrderAttributeToString(pAttribute) & " attribute is not supported by the service provider"

Select Case mStatus
Case OrderStatusCreated
    IsAttributeModifiable = True
Case OrderStatusRejected
    IsAttributeModifiable = True
Case OrderStatusPendingSubmit
    IsAttributeModifiable = attributeId And _
                            (OrderAttributes.OrderAttAllOrNone Or _
                            OrderAttributes.OrderAttBlockOrder Or _
                            OrderAttributes.OrderAttDiscretionaryAmount Or _
                            OrderAttributes.OrderAttDisplaySize Or _
                            OrderAttributes.OrderAttETradeOnly Or _
                            OrderAttributes.OrderAttFirmQuoteOnly Or _
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
Case OrderStatusPreSubmitted
    IsAttributeModifiable = attributeId And _
                            (OrderAttributes.OrderAttAllOrNone Or _
                            OrderAttributes.OrderAttBlockOrder Or _
                            OrderAttributes.OrderAttDiscretionaryAmount Or _
                            OrderAttributes.OrderAttDisplaySize Or _
                            OrderAttributes.OrderAttETradeOnly Or _
                            OrderAttributes.OrderAttFirmQuoteOnly Or _
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
    IsAttributeModifiable = attributeId And _
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

Select Case attributeId
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub SaveRecoveryInfo(pKey As String)
Dim summary As OrderSummary

Const ProcName As String = "SaveRecoveryInfo"

On Error GoTo Err

If mIsSimulated Then Exit Sub
If mStatus = OrderStatuses.OrderStatusCreated Then Exit Sub

Set summary = New OrderSummary

summary.Action = mAction
summary.AllOrNone = mAllOrNone
summary.AveragePrice = mAveragePrice
summary.BlockOrder = False
summary.BrokerId = mBrokerId
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
summary.OriginatorRef = moriginatorRef
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

gTB.BracketOrderRecoveryController.SaveRecoveryInfo pKey, summary

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

''
' Synchronises the specified order to this order so that both are
' identical.
'
' @param  targetOrder the <code>order</code> that must be made identical
'@/
Friend Sub syncToMe(ByVal TargetOrder As Order)
Const ProcName As String = "syncToMe"

On Error GoTo Err

With TargetOrder
    .OrderContext = mOrderContext   ' set this first because others depend on it
    .Status = mStatus               ' ditto
    
    .AveragePrice = mAveragePrice
    .Action = mAction
    If IsAttributeModifiable(OrderAttributes.OrderAttAllOrNone) Then .AllOrNone = mAllOrNone
    If IsAttributeModifiable(OrderAttributes.OrderAttBlockOrder) Then .BlockOrder = mBlockOrder
    .BrokerId = mBrokerId
    If IsAttributeModifiable(OrderAttributes.OrderAttDiscretionaryAmount) Then .DiscretionaryAmount = mDiscretionaryAmount
    If IsAttributeModifiable(OrderAttributes.OrderAttDisplaySize) Then .DisplaySize = mDisplaySize
    .IsSimulated = mIsSimulated
    If IsAttributeModifiable(OrderAttributes.OrderAttETradeOnly) Then .ETradeOnly = mETradeOnly
    .FillTime = mFillTime
    If IsAttributeModifiable(OrderAttributes.OrderAttFirmQuoteOnly) Then .FirmQuoteOnly = mFirmQuoteOnly
    If IsAttributeModifiable(OrderAttributes.OrderAttGoodAfterTime) Then .GoodAfterTime = mGoodAfterTime
    If IsAttributeModifiable(OrderAttributes.OrderAttGoodAfterTimeTZ) Then .GoodAfterTimeTZ = mGoodAfterTimeTZ
    If IsAttributeModifiable(OrderAttributes.OrderAttGoodTillDate) Then .GoodTillDate = mGoodTillDate
    If IsAttributeModifiable(OrderAttributes.OrderAttGoodTillDateTZ) Then .GoodTillDateTZ = mGoodTillDateTZ
    If IsAttributeModifiable(OrderAttributes.OrderAttHidden) Then .Hidden = mHidden
    .Id = mIdentifier
    If IsAttributeModifiable(OrderAttributes.OrderAttIgnoreRTH) Then .IgnoreRegularTradingHours = mIgnoreRegularTradingHours
    .LastFillPrice = mLastFillPrice
    If IsAttributeModifiable(OrderAttributes.OrderAttLimitPrice) Then .LimitPrice = mLimitPrice
    If IsAttributeModifiable(OrderAttributes.OrderAttMinimumQuantity) Then .MinimumQuantity = mMinimumQuantity
    If IsAttributeModifiable(OrderAttributes.OrderAttNBBOPriceCap) Then .NbboPriceCap = mNBBOPriceCap
    .Offset = mOffset
    If IsAttributeModifiable(OrderAttributes.OrderAttOrderType) Then .OrderType = mOrderType
    If IsAttributeModifiable(OrderAttributes.OrderAttOrigin) Then .Origin = mOrigin
    If IsAttributeModifiable(OrderAttributes.OrderAttOriginatorRef) Then .OriginatorRef = moriginatorRef
    If IsAttributeModifiable(OrderAttributes.OrderAttOverrideConstraints) Then .OverrideConstraints = mOverrideConstraints
    If IsAttributeModifiable(OrderAttributes.OrderAttPercentOffset) Then .PercentOffset = mPercentOffset
    If IsAttributeModifiable(OrderAttributes.OrderAttQuantity) Then .Quantity = mQuantity
    .QuantityFilled = mQuantityFilled
    If IsAttributeModifiable(OrderAttributes.OrderAttSettlingFirm) Then .SettlingFirm = mSettlingFirm
    If IsAttributeModifiable(OrderAttributes.OrderAttStopTriggerMethod) And _
        mStopTriggerMethod <> StopTriggerMethods.StopTriggerNone Then .StopTriggerMethod = mStopTriggerMethod
    If IsAttributeModifiable(OrderAttributes.OrderAttSweepToFill) Then .SweepToFill = mSweepToFill
    .Ticker = mDataSource
    If IsAttributeModifiable(OrderAttributes.OrderAttTimeInForce) Then .TimeInForce = mTimeInForce
    If IsAttributeModifiable(OrderAttributes.OrderAttTriggerPrice) Then .TriggerPrice = mTriggerPrice
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub SyncToOrderReport( _
                ByVal OrderReport As IOrderReport, _
                ByVal persistenceKey As String)
Dim changed As Boolean

Const ProcName As String = "SyncToOrderReport"

On Error GoTo Err

With OrderReport
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
    If .OriginatorRef <> moriginatorRef Then moriginatorRef = .OriginatorRef: changed = True
    If .OverrideConstraints <> mOverrideConstraints Then mOverrideConstraints = .OverrideConstraints: changed = True
    If Not propertiesEqual(mProviderProperties, .ProviderProperties) Then Set mProviderProperties = .ProviderProperties: changed = True
    If .Quantity <> mQuantity Then mQuantity = .Quantity: changed = True
    If .SettlingFirm <> mSettlingFirm Then mSettlingFirm = .SettlingFirm: changed = True
    If .StopTriggerMethod <> mStopTriggerMethod Then mStopTriggerMethod = .StopTriggerMethod: changed = True
    If .SweepToFill <> mSweepToFill Then mSweepToFill = .SweepToFill: changed = True
    If .TimeInForce <> mTimeInForce Then mTimeInForce = .TimeInForce: changed = True
    'If .tradebuildId <> mIdentifier Then mIdentifier = .tradebuildId: changed = True
    If .TriggerPrice <> mTriggerPrice Then mTriggerPrice = .TriggerPrice: changed = True
End With

If changed Then
    SaveRecoveryInfo persistenceKey
    RaiseEvent PropertyChanged
End If
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
    If Not IsActive Then NeedsRecovery = False
    SaveRecoveryInfo persistenceKey
    RaiseEvent StatusChanged
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attributeValueNotSupported( _
                ByVal pAttribute As OrderAttributes, _
                ByVal value As Variant)
Const ProcName As String = "attributeValueNotSupported"

On Error GoTo Err

Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value " & CStr(value) & " for attribute " & GOrderAttributeToString(pAttribute) & " is not supported by the service provider"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub checkAttributeModificationPermitted( _
                ByVal attributeId As OrderAttributes)
Const ProcName As String = "checkAttributeModificationPermitted"

On Error GoTo Err

If Not IsAttributeModifiable(attributeId) Then Err.Raise ErrorCodes.ErrIllegalStateException, , "The " & GOrderAttributeToString(attributeId) & " attribute is not currently modifiable"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub





