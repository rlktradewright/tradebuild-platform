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

Private Property Get Action() As TradeBuildSP.OrderActions
End Property

Private Property Get AllOrNone() As Boolean
End Property

Private Property Get BlockOrder() As Boolean
End Property

Private Property Let BrokerId(ByVal RHS As String)
End Property

Private Property Get BrokerId() As String
End Property

Private Property Get Contract() As Contract
End Property

Private Property Get DiscretionaryAmount() As Double
End Property

Private Property Get DisplaySize() As Long
End Property

Private Property Get ETradeOnly() As Boolean
End Property

Private Property Get FirmQuoteOnly() As Boolean
End Property

Private Property Get GoodAfterTime() As Date
End Property

Private Property Get GoodAfterTimeTZ() As String
End Property

Private Property Get GoodTillDate() As Date
End Property

Private Property Get GoodTillDateTZ() As String
End Property

Private Property Get Hidden() As Boolean
End Property

Private Property Get IgnoreRegularTradingHours() As Boolean
End Property

Private Property Get LimitPrice() As Double
End Property

Private Property Get MinimumQuantity() As Long
End Property

Private Property Get NbboPriceCap() As Double
End Property

Private Property Get OrderType() As TradeBuildSP.OrderTypes
End Property

Private Property Get Origin() As TradeBuildSP.OrderOrigins
End Property

Private Property Get OriginatorRef() As String
End Property

Private Property Get OverrideConstraints() As Boolean
End Property

Private Property Get ProviderProperties() As TWUtilities30.Parameters
End Property

Private Property Get Quantity() As Long
End Property

Private Property Get QuantityFilled() As Long
End Property

Private Property Get QuantityRemaining() As Long
End Property

Private Property Get SettlingFirm() As String
End Property

Private Property Get Status() As TradeBuildSP.OrderStatuses
End Property

Private Property Get SweepToFill() As Boolean
End Property

Private Property Get TimeInForce() As TradeBuildSP.OrderTifs
End Property

Private Property Get TradeBuildId() As String
End Property

Private Property Get StopTriggerMethod() As TradeBuildSP.StopTriggerMethods
End Property

Private Property Get TriggerPrice() As Double
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




