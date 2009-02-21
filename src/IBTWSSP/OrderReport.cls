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

Private Const ProjectName                   As String = "IBTWSSP26"
Private Const ModuleName                    As String = "OrderReport"

'@================================================================================
' Member variables
'@================================================================================

Private mAction As TradeBuildSP.OrderActions
Private mallOrNone As Boolean
Private mblockOrder As Boolean
Private mbrokerId As String
Private mContractSpecifier As ContractUtils26.ContractSpecifier
Private mdiscretionaryAmount As Double
Private mdisplaySize As Long
Private meTradeOnly As Boolean
Private mfirmQuoteOnly As Boolean
Private mgoodAfterTime As Date
Private mgoodAfterTimeTZ As String
Private mgoodTillDate As Date
Private mgoodTillDateTZ As String
Private mhidden As Boolean
Private mignoreRegularTradingHours As Boolean
Private mlimitPrice As Double
Private mminimumQuantity As Long
Private mNbboPriceCap As Double
Private morderType As TradeBuildSP.OrderTypes
Private morigin As TradeBuildSP.OrderOrigins
Private moriginatorRef As String
Private moverrideConstraints As Boolean
Private mProviderProperties As Parameters
Private mquantity As Long
Private msettlingFirm As String
Private mStatus As TradeBuildSP.OrderStatuses
Private mstopTriggerMethod As TradeBuildSP.StopTriggerMethods
Private msweepToFill As Boolean
Private mtimeInForce As TradeBuildSP.OrderTifs
Private mtradeBuildId As String
Private mtriggerPrice As Double
 
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

Private Property Let IOrderReport_action(ByVal value As TradeBuildSP.OrderActions)
mAction = value
End Property

Private Property Get IOrderReport_action() As TradeBuildSP.OrderActions
IOrderReport_action = mAction
End Property

Private Property Let IOrderReport_allOrNone(ByVal value As Boolean)
mallOrNone = value
End Property

Private Property Get IOrderReport_allOrNone() As Boolean
IOrderReport_allOrNone = mallOrNone
End Property

Private Property Let IOrderReport_blockOrder(ByVal value As Boolean)
mblockOrder = value
End Property

Private Property Get IOrderReport_blockOrder() As Boolean
IOrderReport_blockOrder = mblockOrder
End Property

Private Property Let IOrderReport_brokerId(ByVal value As String)
mbrokerId = value
End Property

Private Property Get IOrderReport_brokerId() As String
IOrderReport_brokerId = mbrokerId
End Property

Private Property Let IOrderReport_ContractSpecifier(ByVal value As ContractUtils26.ContractSpecifier)
Set mContractSpecifier = value
End Property

Private Property Get IOrderReport_ContractSpecifier() As ContractUtils26.ContractSpecifier
Set IOrderReport_ContractSpecifier = mContractSpecifier
End Property

Private Property Let IOrderReport_discretionaryAmount(ByVal value As Double)
mdiscretionaryAmount = value
End Property

Private Property Get IOrderReport_discretionaryAmount() As Double
IOrderReport_discretionaryAmount = mdiscretionaryAmount
End Property

Private Property Let IOrderReport_displaySize(ByVal value As Long)
mdisplaySize = value
End Property

Private Property Get IOrderReport_displaySize() As Long
IOrderReport_displaySize = mdisplaySize
End Property

Private Property Let IOrderReport_eTradeOnly(ByVal value As Boolean)
meTradeOnly = value
End Property

Private Property Get IOrderReport_eTradeOnly() As Boolean
IOrderReport_eTradeOnly = meTradeOnly
End Property

Private Property Let IOrderReport_firmQuoteOnly(ByVal value As Boolean)
mfirmQuoteOnly = value
End Property

Private Property Get IOrderReport_firmQuoteOnly() As Boolean
IOrderReport_firmQuoteOnly = mfirmQuoteOnly
End Property

Private Property Let IOrderReport_goodAfterTime(ByVal value As Date)
mgoodAfterTime = value
End Property

Private Property Get IOrderReport_goodAfterTime() As Date
IOrderReport_goodAfterTime = mgoodAfterTime
End Property

Private Property Let IOrderReport_goodAfterTimeTZ(ByVal value As String)
mgoodAfterTimeTZ = value
End Property

Private Property Get IOrderReport_goodAfterTimeTZ() As String
IOrderReport_goodAfterTimeTZ = mgoodAfterTimeTZ
End Property

Private Property Let IOrderReport_goodTillDate(ByVal value As Date)
mgoodTillDate = value
End Property

Private Property Get IOrderReport_goodTillDate() As Date
IOrderReport_goodTillDate = mgoodTillDate
End Property

Private Property Let IOrderReport_goodTillDateTZ(ByVal value As String)
mgoodTillDateTZ = value
End Property

Private Property Get IOrderReport_goodTillDateTZ() As String
IOrderReport_goodTillDateTZ = mgoodTillDateTZ
End Property

Private Property Let IOrderReport_hidden(ByVal value As Boolean)
mhidden = value
End Property

Private Property Get IOrderReport_hidden() As Boolean
IOrderReport_hidden = mhidden
End Property

Private Property Let IOrderReport_ignoreRegularTradingHours(ByVal value As Boolean)
mignoreRegularTradingHours = value
End Property

Private Property Get IOrderReport_ignoreRegularTradingHours() As Boolean
IOrderReport_ignoreRegularTradingHours = mignoreRegularTradingHours
End Property

Private Property Let IOrderReport_limitPrice(ByVal value As Double)
mlimitPrice = value
End Property

Private Property Get IOrderReport_limitPrice() As Double
IOrderReport_limitPrice = mlimitPrice
End Property

Private Property Let IOrderReport_minimumQuantity(ByVal value As Long)
mminimumQuantity = value
End Property

Private Property Get IOrderReport_minimumQuantity() As Long
IOrderReport_minimumQuantity = mminimumQuantity
End Property

Private Property Let IOrderReport_NbboPriceCap(ByVal value As Double)
mNbboPriceCap = value
End Property

Private Property Get IOrderReport_NbboPriceCap() As Double
IOrderReport_NbboPriceCap = mNbboPriceCap
End Property

Private Property Let IOrderReport_orderType(ByVal value As TradeBuildSP.OrderTypes)
morderType = value
End Property

Private Property Get IOrderReport_orderType() As TradeBuildSP.OrderTypes
IOrderReport_orderType = morderType
End Property

Private Property Let IOrderReport_origin(ByVal value As TradeBuildSP.OrderOrigins)
morigin = value
End Property

Private Property Get IOrderReport_origin() As TradeBuildSP.OrderOrigins
IOrderReport_origin = morigin
End Property

Private Property Let IOrderReport_originatorRef(ByVal value As String)
moriginatorRef = value
End Property

Private Property Get IOrderReport_originatorRef() As String
IOrderReport_originatorRef = moriginatorRef
End Property

Private Property Let IOrderReport_overrideConstraints(ByVal value As Boolean)
moverrideConstraints = value
End Property

Private Property Get IOrderReport_overrideConstraints() As Boolean
IOrderReport_overrideConstraints = moverrideConstraints
End Property

Private Property Let IOrderReport_providerProperties(ByVal value As TWUtilities30.Parameters)
Set mProviderProperties = value
End Property

Private Property Get IOrderReport_providerProperties() As TWUtilities30.Parameters
Set IOrderReport_providerProperties = mProviderProperties
End Property

Private Property Let IOrderReport_quantity(ByVal value As Long)
mquantity = value
End Property

Private Property Get IOrderReport_quantity() As Long
IOrderReport_quantity = mquantity
End Property

Private Property Let IOrderReport_settlingFirm(ByVal value As String)
msettlingFirm = value
End Property

Private Property Get IOrderReport_settlingFirm() As String
IOrderReport_settlingFirm = msettlingFirm
End Property

Private Property Let IOrderReport_status(ByVal value As TradeBuildSP.OrderStatuses)
mStatus = value
End Property

Private Property Get IOrderReport_status() As TradeBuildSP.OrderStatuses
IOrderReport_status = mStatus
End Property

Private Property Let IOrderReport_stopTriggerMethod(ByVal value As TradeBuildSP.StopTriggerMethods)
mstopTriggerMethod = value
End Property

Private Property Get IOrderReport_stopTriggerMethod() As TradeBuildSP.StopTriggerMethods
IOrderReport_stopTriggerMethod = mstopTriggerMethod
End Property

Private Property Let IOrderReport_sweepToFill(ByVal value As Boolean)
msweepToFill = value
End Property

Private Property Get IOrderReport_sweepToFill() As Boolean
IOrderReport_sweepToFill = msweepToFill
End Property

Private Property Let IOrderReport_timeInForce(ByVal value As TradeBuildSP.OrderTifs)
mtimeInForce = value
End Property

Private Property Get IOrderReport_timeInForce() As TradeBuildSP.OrderTifs
IOrderReport_timeInForce = mtimeInForce
End Property

Private Property Let IOrderReport_tradeBuildId(ByVal value As String)
mtradeBuildId = value
End Property

Private Property Get IOrderReport_tradeBuildId() As String
IOrderReport_tradeBuildId = mtradeBuildId
End Property

Private Property Let IOrderReport_triggerPrice(ByVal value As Double)
mtriggerPrice = value
End Property

Private Property Get IOrderReport_triggerPrice() As Double
IOrderReport_triggerPrice = mtriggerPrice
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


