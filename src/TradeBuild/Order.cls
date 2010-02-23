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

Implements TradeBuildSP.IOrder

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

Private mTicker                     As Ticker
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

Private mTimeInForce                As OrderTifs
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
    
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mStatus = OrderStatuses.OrderStatusCreated
mAction = ActionBuy
mOrderType = OrderTypeNone
mTimeInForce = TIFDay

Set mProviderProperties = New Parameters

End Sub

'@================================================================================
' IOrder Interface Members
'@================================================================================

Private Property Get IOrder_Action() As TradeBuildSP.OrderActions
IOrder_Action = mAction
End Property

Private Property Get IOrder_allOrNone() As Boolean
IOrder_allOrNone = mAllOrNone
End Property

Private Property Get IOrder_blockOrder() As Boolean
IOrder_blockOrder = mBlockOrder
End Property

Private Property Let IOrder_brokerId(ByVal RHS As String)
BrokerId = RHS
End Property

Private Property Get IOrder_brokerId() As String
IOrder_brokerId = mBrokerId
End Property

Private Property Get IOrder_Contract() As Contract
Set IOrder_Contract = mContract
End Property

Private Property Get IOrder_discretionaryAmount() As Double
IOrder_discretionaryAmount = mDiscretionaryAmount
End Property

Private Property Get IOrder_displaySize() As Long
IOrder_displaySize = mDisplaySize
End Property

Private Property Get IOrder_eTradeOnly() As Boolean
IOrder_eTradeOnly = mETradeOnly
End Property

Private Property Get IOrder_firmQuoteOnly() As Boolean
IOrder_firmQuoteOnly = mFirmQuoteOnly
End Property

Private Property Get IOrder_goodAfterTime() As Date
IOrder_goodAfterTime = mGoodAfterTime
End Property

Private Property Get IOrder_goodAfterTimeTZ() As String
IOrder_goodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Private Property Get IOrder_goodTillDate() As Date
IOrder_goodTillDate = mGoodTillDate
End Property

Private Property Get IOrder_goodTillDateTZ() As String
IOrder_goodTillDateTZ = mGoodTillDateTZ
End Property

Private Property Get IOrder_hidden() As Boolean
IOrder_hidden = mHidden
End Property

Private Property Get IOrder_ignoreRegularTradingHours() As Boolean
IOrder_ignoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Private Property Get IOrder_limitPrice() As Double
IOrder_limitPrice = mLimitPrice
End Property

Private Property Get IOrder_minimumQuantity() As Long
IOrder_minimumQuantity = mMinimumQuantity
End Property

Private Property Get IOrder_NbboPriceCap() As Double
IOrder_NbboPriceCap = mNBBOPriceCap
End Property

Private Property Get IOrder_orderType() As TradeBuildSP.OrderTypes
IOrder_orderType = mOrderType
End Property

Private Property Get IOrder_origin() As TradeBuildSP.OrderOrigins
IOrder_origin = mOrigin
End Property

Private Property Get IOrder_originatorRef() As String
IOrder_originatorRef = moriginatorRef
End Property

Private Property Get IOrder_overrideConstraints() As Boolean
IOrder_overrideConstraints = mOverrideConstraints
End Property

Private Property Get IOrder_providerProperties() As TWUtilities30.Parameters
Set IOrder_providerProperties = mProviderProperties
End Property

Private Property Get IOrder_quantity() As Long
IOrder_quantity = mQuantity
End Property

Private Property Get IOrder_quantityFilled() As Long
IOrder_quantityFilled = mQuantityFilled
End Property

Private Property Get IOrder_quantityRemaining() As Long
IOrder_quantityRemaining = mQuantityRemaining
End Property

Private Property Get IOrder_settlingFirm() As String
IOrder_settlingFirm = mSettlingFirm
End Property

Private Property Get IOrder_status() As TradeBuildSP.OrderStatuses
IOrder_status = mStatus
End Property

Private Property Get IOrder_sweepToFill() As Boolean
IOrder_sweepToFill = mSweepToFill
End Property

Private Property Get IOrder_timeInForce() As TradeBuildSP.OrderTifs
IOrder_timeInForce = mTimeInForce
End Property

Private Property Get IOrder_tradeBuildId() As String
IOrder_tradeBuildId = mIdentifier
End Property

Private Property Get IOrder_stopTriggerMethod() As TradeBuildSP.StopTriggerMethods
IOrder_stopTriggerMethod = mStopTriggerMethod
End Property

Private Property Get IOrder_triggerPrice() As Double
IOrder_triggerPrice = mTriggerPrice
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let action(ByVal value As OrderActions)
mAction = value
End Property

Public Property Get action() As OrderActions
action = mAction
End Property

Friend Property Get Active() As Boolean
Active = (mStatus = OrderStatusPendingSubmit) Or _
        (mStatus = OrderStatusPreSubmitted) Or _
        (mStatus = OrderStatusSubmitted)
End Property

Public Property Let AllOrNone(ByVal value As Boolean)
Const ProcName As String = "allOrNone"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttAllOrNone
If value = mAllOrNone Then Exit Property
setDirty
mAllOrNone = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "blockOrder"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttBlockOrder
If value = mBlockOrder Then Exit Property
setDirty
mBlockOrder = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Public Property Let DiscretionaryAmount(ByVal value As Double)
Const ProcName As String = "discretionaryAmount"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttDiscretionaryAmount
If value = mDiscretionaryAmount Then Exit Property
setDirty
mDiscretionaryAmount = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DiscretionaryAmount() As Double
DiscretionaryAmount = mDiscretionaryAmount
End Property

Public Property Let DisplaySize(ByVal value As Long)
Const ProcName As String = "displaySize"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttDisplaySize
If value = mDisplaySize Then Exit Property
setDirty
mDisplaySize = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mDisplaySize
End Property

Public Property Let ETradeOnly(ByVal value As Boolean)
Const ProcName As String = "eTradeOnly"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttETradeOnly
If value = mETradeOnly Then Exit Property
setDirty
mETradeOnly = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "firmQuoteOnly"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttFirmQuoteOnly
If value = mFirmQuoteOnly Then Exit Property
setDirty
mFirmQuoteOnly = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get FirmQuoteOnly() As Boolean
FirmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let GoodAfterTime(ByVal value As Date)
Const ProcName As String = "goodAfterTime"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttGoodAfterTime
If value = mGoodAfterTime Then Exit Property
setDirty
mGoodAfterTime = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get GoodAfterTime() As Date
GoodAfterTime = mGoodAfterTime
End Property

Public Property Let GoodAfterTimeTZ(ByVal value As String)
Const ProcName As String = "goodAfterTimeTZ"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttGoodAfterTimeTZ
If value = mGoodAfterTimeTZ Then Exit Property
setDirty
mGoodAfterTimeTZ = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get GoodAfterTimeTZ() As String
GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let GoodTillDate(ByVal value As Date)
Const ProcName As String = "goodTillDate"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttGoodTillDate
If value = mGoodTillDate Then Exit Property
setDirty
mGoodTillDate = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get GoodTillDate() As Date
GoodTillDate = mGoodTillDate
End Property

Public Property Let GoodTillDateTZ(ByVal value As String)
Const ProcName As String = "goodTillDateTZ"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttGoodTillDateTZ
If value = mGoodTillDateTZ Then Exit Property
setDirty
mGoodTillDateTZ = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get GoodTillDateTZ() As String
GoodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let Hidden(value As Boolean)
Const ProcName As String = "hidden"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttHidden
If value = mHidden Then Exit Property
setDirty
mHidden = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "ignoreRegularTradingHours"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttIgnoreRTH
If value = mIgnoreRegularTradingHours Then Exit Property
setDirty
mIgnoreRegularTradingHours = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IgnoreRegularTradingHours() As Boolean
IgnoreRegularTradingHours = mIgnoreRegularTradingHours
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
Const ProcName As String = "minimumQuantity"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttMinimumQuantity
If value = mMinimumQuantity Then Exit Property
setDirty
mMinimumQuantity = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get MinimumQuantity() As Long
MinimumQuantity = mMinimumQuantity
End Property

Public Property Let NbboPriceCap(ByVal value As Double)
Const ProcName As String = "NbboPriceCap"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttNBBOPriceCap
If value = mNBBOPriceCap Then Exit Property
setDirty
mNBBOPriceCap = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
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
Const ProcName As String = "orderType"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttOrderType
If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsOrderTypeValueSupported(value) Then attributeValueNotSupported TradeBuild26.OrderAttributeIds.OrderAttOrderType, value
End If
If value = mOrderType Then Exit Property
setDirty
mOrderType = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mOrderType
End Property

Public Property Let Origin(ByVal value As OrderOrigins)
Const ProcName As String = "origin"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttOrigin
If value = mOrigin Then Exit Property
setDirty
mOrigin = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Origin() As OrderOrigins
Origin = mOrigin
End Property

Public Property Let OriginatorRef(ByVal value As String)
Const ProcName As String = "originatorRef"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttOriginatorRef
If value = moriginatorRef Then Exit Property
setDirty
moriginatorRef = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get OriginatorRef() As String
OriginatorRef = moriginatorRef
End Property

Public Property Let OverrideConstraints(ByVal value As Boolean)
Const ProcName As String = "overrideConstraints"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttOverrideConstraints
If value = mOverrideConstraints Then Exit Property
setDirty
mOverrideConstraints = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get OverrideConstraints() As Boolean
OverrideConstraints = mOverrideConstraints
End Property

Public Property Let PercentOffset(ByVal value As Double)
Const ProcName As String = "percentOffset"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttPercentOffset
If value = mPercentOffset Then Exit Property
setDirty
mPercentOffset = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "quantity"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttQuantity
If value = mQuantity Then Exit Property
mQuantity = value
setDirty
mQuantityRemaining = mQuantity - mQuantityFilled

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "settlingFirm"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttSettlingFirm
If value = mSettlingFirm Then Exit Property
setDirty
mSettlingFirm = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttStopTriggerMethod
If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsStopTriggerMethodValueSupported(value) Then attributeValueNotSupported TradeBuild26.OrderAttributeIds.OrderAttStopTriggerMethod, value
End If
If value = mStopTriggerMethod Then Exit Property
setDirty
mStopTriggerMethod = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get StopTriggerMethod() As StopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(value As Boolean)
Const ProcName As String = "SweepToFill"
Dim failpoint As Long
On Error GoTo Err

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttSweepToFill
If value = mSweepToFill Then Exit Property
setDirty
mSweepToFill = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Const ProcName As String = "Ticker"
Dim failpoint As Long
On Error GoTo Err

Set mTicker = value
Set mContract = mTicker.Contract

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

Public Property Let TimeInForce(ByVal value As OrderTifs)
Const ProcName As String = "timeInForce"
Dim failpoint As Long
On Error GoTo Err

If Not GIsValidTIF(value) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must be a member of the OrderTifs enum"
End If

checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttTimeInForce
If Not mOrderContext Is Nothing Then
    If Not mOrderContext.IsOrderTifValueSupported(value) Then attributeValueNotSupported TradeBuild26.OrderAttributeIds.OrderAttTimeInForce, value
End If
If value = mTimeInForce Then Exit Property
setDirty
mTimeInForce = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TimeInForce() As OrderTifs
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
Const ProcName As String = "clone"
Dim failpoint As Long
On Error GoTo Err

Set Clone = New Order
syncToMe Clone

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

''
' Notifies a fill for this <code>order</code>.
'
' @param exec  an <code>Execution</code> object containing details of the fill
'
'@/
Friend Sub NotifyFill( _
                ByVal exec As Execution, _
                ByRef persistenceKey As String)

Const ProcName As String = "NotifyFill"
Dim failpoint As Long
On Error GoTo Err

If mContract Is Nothing Then
    ' happens when an order from a previous session is being recovered before
    ' the relevant ticker is ready
Else
    GLogOrderDetail mContract.specifier.Exchange & "," & _
            mContract.specifier.Symbol & "," & _
            mContract.specifier.Expiry & "," & _
            exec.Quantity & "," & _
            exec.price & "," & _
            GOrderActionToString(exec.action) & "," & _
            exec.Time & "," & _
            exec.Timezone & "," & _
            exec.OrderTradeBuildId & "," & _
            exec.OrderBrokerId, _
            mIsSimulated, _
            Me
End If

mAveragePrice = (mQuantityFilled * mAveragePrice + exec.Quantity * exec.price) / (mQuantityFilled + exec.Quantity)
mQuantityFilled = mQuantityFilled + exec.Quantity
mQuantityRemaining = mQuantityRemaining - exec.Quantity
mLastFillPrice = exec.price
mFillTime = exec.Time
If mQuantityRemaining = 0 Then mStatus = OrderStatusFilled
SaveRecoveryInfo persistenceKey
RaiseEvent Fill(exec)
If mStatus = OrderStatusFilled Then RaiseEvent StatusChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As Long
On Error GoTo Err

Set mTicker = Nothing
Set mContract = Nothing
Set mOrderContext = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function IsAttributeModifiable( _
                ByVal attributeId As OrderAttributeIds) As Boolean

Const ProcName As String = "isAttributeModifiable"
Dim failpoint As Long
On Error GoTo Err

If mOrderContext Is Nothing Then
    ' we are currently recovering persisted order information so don't make this check
    IsAttributeModifiable = True
    Exit Function
End If

If Not mOrderContext.IsAttributeSupported(attributeId) Then Exit Function

Select Case mStatus
Case OrderStatusCreated
    IsAttributeModifiable = True
Case OrderStatusRejected
    IsAttributeModifiable = True
Case OrderStatusPendingSubmit
    IsAttributeModifiable = attributeId And _
                            (OrderAttributeIds.OrderAttAllOrNone Or _
                            OrderAttributeIds.OrderAttBlockOrder Or _
                            OrderAttributeIds.OrderAttDiscretionaryAmount Or _
                            OrderAttributeIds.OrderAttDisplaySize Or _
                            OrderAttributeIds.OrderAttETradeOnly Or _
                            OrderAttributeIds.OrderAttFirmQuoteOnly Or _
                            OrderAttributeIds.OrderAttGoodTillDate Or _
                            OrderAttributeIds.OrderAttGoodTillDateTZ Or _
                            OrderAttributeIds.OrderAttHidden Or _
                            OrderAttributeIds.OrderAttLimitPrice Or _
                            OrderAttributeIds.OrderAttMinimumQuantity Or _
                            OrderAttributeIds.OrderAttOriginatorRef Or _
                            OrderAttributeIds.OrderAttQuantity Or _
                            OrderAttributeIds.OrderAttSweepToFill Or _
                            OrderAttributeIds.OrderAttTimeInForce Or _
                            OrderAttributeIds.OrderAttTriggerPrice)
Case OrderStatusPreSubmitted
    IsAttributeModifiable = attributeId And _
                            (OrderAttributeIds.OrderAttAllOrNone Or _
                            OrderAttributeIds.OrderAttBlockOrder Or _
                            OrderAttributeIds.OrderAttDiscretionaryAmount Or _
                            OrderAttributeIds.OrderAttDisplaySize Or _
                            OrderAttributeIds.OrderAttETradeOnly Or _
                            OrderAttributeIds.OrderAttFirmQuoteOnly Or _
                            OrderAttributeIds.OrderAttGoodTillDate Or _
                            OrderAttributeIds.OrderAttGoodTillDateTZ Or _
                            OrderAttributeIds.OrderAttHidden Or _
                            OrderAttributeIds.OrderAttLimitPrice Or _
                            OrderAttributeIds.OrderAttMinimumQuantity Or _
                            OrderAttributeIds.OrderAttOriginatorRef Or _
                            OrderAttributeIds.OrderAttQuantity Or _
                            OrderAttributeIds.OrderAttSweepToFill Or _
                            OrderAttributeIds.OrderAttTimeInForce Or _
                            OrderAttributeIds.OrderAttTriggerPrice)
Case OrderStatusSubmitted
    IsAttributeModifiable = attributeId And _
                            (OrderAttributeIds.OrderAttDiscretionaryAmount Or _
                            OrderAttributeIds.OrderAttDisplaySize Or _
                            OrderAttributeIds.OrderAttGoodTillDate Or _
                            OrderAttributeIds.OrderAttGoodTillDateTZ Or _
                            OrderAttributeIds.OrderAttHidden Or _
                            OrderAttributeIds.OrderAttLimitPrice Or _
                            OrderAttributeIds.OrderAttMinimumQuantity Or _
                            OrderAttributeIds.OrderAttQuantity Or _
                            OrderAttributeIds.OrderAttTimeInForce Or _
                            OrderAttributeIds.OrderAttTriggerPrice)
Case OrderStatusFilled
    IsAttributeModifiable = False
Case OrderStatusCancelling
    IsAttributeModifiable = False
Case OrderStatusCancelled
    IsAttributeModifiable = False
End Select

Select Case attributeId
Case OrderAttributeIds.OrderAttLimitPrice
    Select Case mOrderType
    Case OrderTypeLimit, _
        OrderTypeLimitOnClose, _
        OrderTypeStopLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeLimitOnOpen
    Case Else
        IsAttributeModifiable = False
    End Select
Case OrderAttributeIds.OrderAttTriggerPrice
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function IsModifiable() As Boolean
Const ProcName As String = "isModifiable"
Dim failpoint As Long
On Error GoTo Err

If mStatus <> OrderStatuses.OrderStatusCancelled And _
    mStatus <> OrderStatuses.OrderStatusCancelling And _
    mStatus <> OrderStatuses.OrderStatusFilled _
Then IsModifiable = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

' commented out 22/5/6
'Friend Sub openOrder(ByVal Action As OrderActions, _
'                        ByVal quantity As Long, _
'                        ByVal orderType As OrderTypes, _
'                        ByVal lmtPrice As Double, _
'                        ByVal auxPrice As Double, _
'                        ByVal tif As OrderTifs, _
'                        ByVal ocaGroup As String)
'
'mTWSAction = Action
'If Action <> mAction Then
'    mDifferent = True
'End If
'
'mTWSQuantity = quantity
'If quantity <> mQuantity Then
'    mDifferent = True
'End If
'
'mTWSOrderType = orderType
'If orderType <> mOrderType Then
'    mDifferent = True
'End If
'If orderType = OrderTypes.OrderTypeLimit Or _
'    orderType = OrderTypes.OrderTypeStopLimit _
'Then
'    mTWSLimitPrice = lmtPrice
'    If lmtPrice <> mLimitPrice Then
'        mDifferent = True
'    End If
'Else
'    mTWSLimitPrice = 0
'End If
'If orderType = OrderTypes.OrderTypeStop Or _
'    orderType = OrderTypes.OrderTypeStopLimit _
'Then
'    mTWSTriggerPrice = auxPrice
'    If auxPrice <> mTriggerPrice Then
'        mDifferent = True
'    End If
'Else
'    mTWSTriggerPrice = 0
'End If
'' NB: mDifferent will be cleared when UpdateStatus is called (this always
'' happens shortly after openOrder)
''If mDifferent Then Stop
'End Sub

Friend Sub SaveRecoveryInfo(pKey As String)
Dim summary As OrderSummary

Const ProcName As String = "SaveRecoveryInfo"
Dim failpoint As String
On Error GoTo Err

If mIsSimulated Then Exit Sub
If mStatus = OrderStatuses.OrderStatusCreated Then Exit Sub

Set summary = New OrderSummary

summary.action = mAction
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

gTB.OrderPlexRecoveryController.SaveRecoveryInfo pKey, summary.ToJSON

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

''
' Marks this order as clean, ie all changes have been either
' actioned or cancelled.
'@/
Friend Sub setClean()
Const ProcName As String = "setClean"
Dim failpoint As Long
On Error GoTo Err

mIsDirty = False
RaiseEvent Clean

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

''
' Synchronises the specified order to this order so that both are
' identical.
'
' @param  targetOrder the <code>order</code> that must be made identical
'@/
Friend Sub syncToMe(ByVal TargetOrder As Order)
Const ProcName As String = "syncToMe"
Dim failpoint As Long
On Error GoTo Err

With TargetOrder
    .OrderContext = mOrderContext   ' set this first because others depend on it
    .Status = mStatus               ' ditto
    
    .AveragePrice = mAveragePrice
    .action = mAction
    If IsAttributeModifiable(OrderAttributeIds.OrderAttAllOrNone) Then .AllOrNone = mAllOrNone
    If IsAttributeModifiable(OrderAttributeIds.OrderAttBlockOrder) Then .BlockOrder = mBlockOrder
    .BrokerId = mBrokerId
    If IsAttributeModifiable(OrderAttributeIds.OrderAttDiscretionaryAmount) Then .DiscretionaryAmount = mDiscretionaryAmount
    If IsAttributeModifiable(OrderAttributeIds.OrderAttDisplaySize) Then .DisplaySize = mDisplaySize
    .IsSimulated = mIsSimulated
    If IsAttributeModifiable(OrderAttributeIds.OrderAttETradeOnly) Then .ETradeOnly = mETradeOnly
    .FillTime = mFillTime
    If IsAttributeModifiable(OrderAttributeIds.OrderAttFirmQuoteOnly) Then .FirmQuoteOnly = mFirmQuoteOnly
    If IsAttributeModifiable(OrderAttributeIds.OrderAttGoodAfterTime) Then .GoodAfterTime = mGoodAfterTime
    If IsAttributeModifiable(OrderAttributeIds.OrderAttGoodAfterTimeTZ) Then .GoodAfterTimeTZ = mGoodAfterTimeTZ
    If IsAttributeModifiable(OrderAttributeIds.OrderAttGoodTillDate) Then .GoodTillDate = mGoodTillDate
    If IsAttributeModifiable(OrderAttributeIds.OrderAttGoodTillDateTZ) Then .GoodTillDateTZ = mGoodTillDateTZ
    If IsAttributeModifiable(OrderAttributeIds.OrderAttHidden) Then .Hidden = mHidden
    .Id = mIdentifier
    If IsAttributeModifiable(OrderAttributeIds.OrderAttIgnoreRTH) Then .IgnoreRegularTradingHours = mIgnoreRegularTradingHours
    .LastFillPrice = mLastFillPrice
    If IsAttributeModifiable(OrderAttributeIds.OrderAttLimitPrice) Then .LimitPrice = mLimitPrice
    If IsAttributeModifiable(OrderAttributeIds.OrderAttMinimumQuantity) Then .MinimumQuantity = mMinimumQuantity
    If IsAttributeModifiable(OrderAttributeIds.OrderAttNBBOPriceCap) Then .NbboPriceCap = mNBBOPriceCap
    .Offset = mOffset
    If IsAttributeModifiable(OrderAttributeIds.OrderAttOrderType) Then .OrderType = mOrderType
    If IsAttributeModifiable(OrderAttributeIds.OrderAttOrigin) Then .Origin = mOrigin
    If IsAttributeModifiable(OrderAttributeIds.OrderAttOriginatorRef) Then .OriginatorRef = moriginatorRef
    If IsAttributeModifiable(OrderAttributeIds.OrderAttOverrideConstraints) Then .OverrideConstraints = mOverrideConstraints
    If IsAttributeModifiable(OrderAttributeIds.OrderAttPercentOffset) Then .PercentOffset = mPercentOffset
    If IsAttributeModifiable(OrderAttributeIds.OrderAttQuantity) Then .Quantity = mQuantity
    .QuantityFilled = mQuantityFilled
    If IsAttributeModifiable(OrderAttributeIds.OrderAttSettlingFirm) Then .SettlingFirm = mSettlingFirm
    If IsAttributeModifiable(OrderAttributeIds.OrderAttStopTriggerMethod) And _
        mStopTriggerMethod <> StopTriggerMethods.StopTriggerNone Then .StopTriggerMethod = mStopTriggerMethod
    If IsAttributeModifiable(OrderAttributeIds.OrderAttSweepToFill) Then .SweepToFill = mSweepToFill
    .Ticker = mTicker
    If IsAttributeModifiable(OrderAttributeIds.OrderAttTimeInForce) Then .TimeInForce = mTimeInForce
    If IsAttributeModifiable(OrderAttributeIds.OrderAttTriggerPrice) Then .TriggerPrice = mTriggerPrice
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub SyncToOrderReport( _
                ByVal OrderReport As IOrderReport, _
                ByVal persistenceKey As String)
Dim changed As Boolean

Const ProcName As String = "SyncToOrderReport"
Dim failpoint As Long
On Error GoTo Err

With OrderReport
    If .action <> mAction Then mAction = .action: changed = True
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
    Set mProviderProperties = .ProviderProperties
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
                Optional ByRef persistenceKey As String)

Const ProcName As String = "UpdateStatus"
Dim failpoint As Long
On Error GoTo Err

If pStatus <> mStatus Then
    mStatus = pStatus
    SaveRecoveryInfo persistenceKey
    RaiseEvent StatusChanged
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attributeValueNotSupported( _
                ByVal pAttribute As OrderAttributeIds, _
                ByVal value As Variant)
Const ProcName As String = "attributeValueNotSupported"
Dim failpoint As Long
On Error GoTo Err

Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value " & CStr(value) & " for attribute " & GOrderAttributeToString(pAttribute) & " is not supported by the service provider"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub checkAttributeModificationPermitted( _
                ByVal attributeId As OrderAttributeIds)
Const ProcName As String = "checkAttributeModificationPermitted"
Dim failpoint As Long
On Error GoTo Err

If Not IsAttributeModifiable(attributeId) Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "The " & GOrderAttributeToString(attributeId) & " attribute is not currently modifiable"
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

''
' Set the 'dirty' flag, and raise a Dirty event if one has not alReady
' been raised.
'@/
Private Sub setDirty()
Const ProcName As String = "setDirty"
Dim failpoint As Long
On Error GoTo Err

If Not mIsDirty Then
    mIsDirty = True
    RaiseEvent Dirty
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub





