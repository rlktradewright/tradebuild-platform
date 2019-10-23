VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrder"
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

Private Const ModuleName                            As String = "IOrder"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Action(ByVal Value As OrderActions)
End Property

Public Property Get Action() As OrderActions
End Property

Public Property Let AllOrNone(ByVal Value As Boolean)
End Property

Public Property Get AllOrNone() As Boolean
End Property

Public Property Let AveragePrice(ByVal Value As Double)
End Property

Public Property Get AveragePrice() As Double
End Property

Public Property Let BlockOrder(ByVal Value As Boolean)
End Property

Public Property Get BlockOrder() As Boolean
End Property

Public Property Let BrokerId(ByVal Value As String)
End Property

Public Property Get BrokerId() As String
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
End Property

Public Property Let DiscretionaryAmount(ByVal Value As Double)
End Property

Public Property Get DiscretionaryAmount() As Double
End Property

Public Property Let DisplaySize(ByVal Value As Long)
End Property

Public Property Get DisplaySize() As Long
End Property

Public Property Let ErrorCode(ByVal Value As Long)
End Property

Public Property Get ErrorCode() As Long
End Property

Public Property Let ErrorMessage(ByVal Value As String)
End Property

Public Property Get ErrorMessage() As String
End Property

Public Property Let ETradeOnly(ByVal Value As Boolean)
End Property

Public Property Get ETradeOnly() As Boolean
End Property

Public Property Let FillTime(ByVal Value As Date)
End Property

Public Property Get FillTime() As Date
End Property

Public Property Let FirmQuoteOnly(ByVal Value As Boolean)
End Property

Public Property Get FirmQuoteOnly() As Boolean
End Property

Public Property Let GoodAfterTime(ByVal Value As Date)
End Property

Public Property Get GoodAfterTime() As Date
End Property

Public Property Let GoodAfterTimeTZ(ByVal Value As String)
End Property

Public Property Get GoodAfterTimeTZ() As String
End Property

Public Property Let GoodTillDate(ByVal Value As Date)
End Property

Public Property Get GoodTillDate() As Date
End Property

Public Property Let GoodTillDateTZ(ByVal Value As String)
End Property

Public Property Get GoodTillDateTZ() As String
End Property

Public Property Get GroupName() As String
End Property

Public Property Let Hidden(ByVal Value As Boolean)
End Property

Public Property Get Hidden() As Boolean
End Property

Public Property Let Id(ByVal Value As String)
End Property

Public Property Get Id() As String
End Property

Public Property Let IgnoreRegularTradingHours(ByVal Value As Boolean)
End Property

Public Property Get IgnoreRegularTradingHours() As Boolean
End Property

Public Property Get IsActive() As Boolean
End Property

Public Property Get IsDirty() As Boolean
End Property

Public Property Let IsSimulated(ByVal Value As Boolean)
End Property

Public Property Get IsSimulated() As Boolean
End Property

Public Property Let LastFillPrice(ByVal Value As Double)
End Property

Public Property Get LastFillPrice() As Double
End Property

Public Property Let LimitPrice(ByVal Value As Double)
End Property

Public Property Get LimitPrice() As Double
End Property

Public Property Let LimitPriceSpec(ByVal Value As PriceSpecifier)
End Property

Public Property Get LimitPriceSpec() As PriceSpecifier
End Property

Public Property Let Message(ByVal Value As String)
End Property

Public Property Get Message() As String
End Property

Public Property Let MinimumQuantity(ByVal Value As Long)
End Property

Public Property Get MinimumQuantity() As Long
End Property

Public Property Let NbboPriceCap(ByVal Value As Double)
End Property

Public Property Get NbboPriceCap() As Double
End Property

Public Property Get OrderContext() As OrderContext
End Property

Public Property Get OrderType() As OrderTypes
End Property

Public Property Let Origin(ByVal Value As OrderOrigins)
End Property

Public Property Get Origin() As OrderOrigins
End Property

Public Property Let OriginatorRef(ByVal Value As String)
End Property

Public Property Get OriginatorRef() As String
End Property

Public Property Let OverrideConstraints(ByVal Value As Boolean)
End Property

Public Property Get OverrideConstraints() As Boolean
End Property

Public Property Let PercentOffset(ByVal Value As Double)
End Property

Public Property Get PercentOffset() As Double
End Property

Public Property Get ProviderProperties() As Parameters
End Property

Public Property Let Quantity(ByVal Value As Long)
End Property

Public Property Get Quantity() As Long
End Property

Public Property Let QuantityFilled(ByVal Value As Long)
End Property

Public Property Get QuantityFilled() As Long
End Property

Public Property Let QuantityRemaining(ByVal Value As Long)
End Property

Public Property Get QuantityRemaining() As Long
End Property

Public Property Let SettlingFirm(ByVal Value As String)
End Property

Public Property Get SettlingFirm() As String
End Property

Public Property Let Status(ByVal Value As OrderStatuses)
End Property

Public Property Get Status() As OrderStatuses
End Property

Public Property Let StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
End Property

Public Property Get StopTriggerMethod() As OrderStopTriggerMethods
End Property

Public Property Let SweepToFill(ByVal Value As Boolean)
End Property

Public Property Get SweepToFill() As Boolean
End Property

Public Property Let TimeInForce(ByVal Value As OrderTIFs)
End Property

Public Property Get TimeInForce() As OrderTIFs
End Property

Public Property Let TriggerPrice(ByVal Value As Double)
End Property

Public Property Get TriggerPrice() As Double
End Property

Public Property Let TriggerPriceSpec(ByVal Value As PriceSpecifier)
End Property

Public Property Get TriggerPriceSpec() As PriceSpecifier
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pGroupName As String, _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pOrderContext As OrderContext)
End Sub

Public Function IsAttributeModifiable( _
                ByVal pAttribute As OrderAttributes) As Boolean
End Function

Public Function IsModifiable() As Boolean
End Function

Public Sub NotifyError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
End Sub

Public Sub NotifyMessage(ByVal pMessage As String)
End Sub

Public Sub NotifyFill(ByVal pExec As IExecutionReport)
End Sub

Public Sub SyncToOrderReport(ByVal pOrderReport As IOrderReport)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




