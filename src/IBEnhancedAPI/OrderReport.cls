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
Attribute VB_Creatable = False
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
Private mdiscretionaryAmount As Double
Private mDisplaySize As Long
Private mETradeOnly As Boolean
Private mFirmQuoteOnly As Boolean
Private mGoodAfterTime As Date
Private mgoodAfterTimeTZ As String
Private mGoodTillDate As Date
Private mgoodTillDateTZ As String
Private mHidden As Boolean
Private mignoreRegularTradingHours As Boolean
Private mLimitPrice As Double
Private mMinimumQuantity As Long
Private mNbboPriceCap As Double
Private mOrderType As OrderTypes
Private mOrigin As OrderOrigins
Private mOriginatorRef As String
Private mOverrideConstraints As Boolean
Private mProviderProperties As Parameters
Private mQuantity As Long
Private mSettlingFirm As String
Private mStatus As OrderStatuses
Private mStopTriggerMethod As StopTriggerMethods
Private mSweepToFill As Boolean
Private mTimeInForce As OrderTIFs
Private mIdentifier As String
Private mTriggerPrice As Double
 
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mProviderProperties = New Parameters
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
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
mdiscretionaryAmount = Value
End Property

Private Property Get IOrderReport_DiscretionaryAmount() As Double
IOrderReport_DiscretionaryAmount = mdiscretionaryAmount
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
mgoodAfterTimeTZ = Value
End Property

Private Property Get IOrderReport_GoodAfterTimeTZ() As String
IOrderReport_GoodAfterTimeTZ = mgoodAfterTimeTZ
End Property

Private Property Let IOrderReport_GoodTillDate(ByVal Value As Date)
mGoodTillDate = Value
End Property

Private Property Get IOrderReport_GoodTillDate() As Date
IOrderReport_GoodTillDate = mGoodTillDate
End Property

Private Property Let IOrderReport_GoodTillDateTZ(ByVal Value As String)
mgoodTillDateTZ = Value
End Property

Private Property Get IOrderReport_GoodTillDateTZ() As String
IOrderReport_GoodTillDateTZ = mgoodTillDateTZ
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
mignoreRegularTradingHours = Value
End Property

Private Property Get IOrderReport_IgnoreRegularTradingHours() As Boolean
IOrderReport_IgnoreRegularTradingHours = mignoreRegularTradingHours
End Property

Private Property Let IOrderReport_LimitPrice(ByVal Value As Double)
mLimitPrice = Value
End Property

Private Property Get IOrderReport_LimitPrice() As Double
IOrderReport_LimitPrice = mLimitPrice
End Property

Private Property Let IOrderReport_MinimumQuantity(ByVal Value As Long)
mMinimumQuantity = Value
End Property

Private Property Get IOrderReport_MinimumQuantity() As Long
IOrderReport_MinimumQuantity = mMinimumQuantity
End Property

Private Property Let IOrderReport_NbboPriceCap(ByVal Value As Double)
mNbboPriceCap = Value
End Property

Private Property Get IOrderReport_NbboPriceCap() As Double
IOrderReport_NbboPriceCap = mNbboPriceCap
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

Private Property Let IOrderReport_ProviderProperties(ByVal Value As Parameters)
Set mProviderProperties = Value
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

Private Property Let IOrderReport_StopTriggerMethod(ByVal Value As StopTriggerMethods)
mStopTriggerMethod = Value
End Property

Private Property Get IOrderReport_StopTriggerMethod() As StopTriggerMethods
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
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


