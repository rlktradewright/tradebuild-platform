VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrderReport"
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

Private Const ModuleName                            As String = "IOrderReport"

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

Public Property Let BlockOrder(ByVal Value As Boolean)
End Property

Public Property Get BlockOrder() As Boolean
End Property

Public Property Let BrokerId(ByVal Value As String)
End Property

Public Property Get BrokerId() As String
End Property

Public Property Let DiscretionaryAmount(ByVal Value As Double)
End Property

Public Property Get DiscretionaryAmount() As Double
End Property

Public Property Let DisplaySize(ByVal Value As Long)
End Property

Public Property Get DisplaySize() As Long
End Property

Public Property Let ETradeOnly(ByVal Value As Boolean)
End Property

Public Property Get ETradeOnly() As Boolean
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

Public Property Let IsSimulated(ByVal Value As Boolean)
End Property

Public Property Get IsSimulated() As Boolean
End Property

Public Property Let LimitPrice(ByVal Value As Double)
End Property

Public Property Get LimitPrice() As Double
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

Public Property Let OrderType(ByVal Value As OrderTypes)
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

Public Property Get ProviderProperties() As Parameters
End Property

Public Property Let Quantity(ByVal Value As Long)
End Property

Public Property Get Quantity() As Long
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

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




