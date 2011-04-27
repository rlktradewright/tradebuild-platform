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

Private Property Let action(ByVal Value As OrderActions)
End Property

Private Property Get action() As OrderActions
End Property

Private Property Let allOrNone(ByVal Value As Boolean)
End Property

Private Property Get allOrNone() As Boolean
End Property

Private Property Let blockOrder(ByVal Value As Boolean)
End Property

Private Property Get blockOrder() As Boolean
End Property

Private Property Let brokerId(ByVal Value As String)
End Property

Private Property Get brokerId() As String
End Property

Private Property Let ContractSpecifier(ByVal Value As ContractUtils27.ContractSpecifier)
End Property

Private Property Get ContractSpecifier() As ContractUtils27.ContractSpecifier
End Property

Private Property Let discretionaryAmount(ByVal Value As Double)
End Property

Private Property Get discretionaryAmount() As Double
End Property

Private Property Let displaySize(ByVal Value As Long)
End Property

Private Property Get displaySize() As Long
End Property

Private Property Let eTradeOnly(ByVal Value As Boolean)
End Property

Private Property Get eTradeOnly() As Boolean
End Property

Private Property Let firmQuoteOnly(ByVal Value As Boolean)
End Property

Private Property Get firmQuoteOnly() As Boolean
End Property

Private Property Let goodAfterTime(ByVal Value As Date)
End Property

Private Property Get goodAfterTime() As Date
End Property

Private Property Let goodAfterTimeTZ(ByVal Value As String)
End Property

Private Property Get goodAfterTimeTZ() As String
End Property

Private Property Let goodTillDate(ByVal Value As Date)
End Property

Private Property Get goodTillDate() As Date
End Property

Private Property Let goodTillDateTZ(ByVal Value As String)
End Property

Private Property Get goodTillDateTZ() As String
End Property

Private Property Let hidden(ByVal Value As Boolean)
End Property

Private Property Get hidden() As Boolean
End Property

Private Property Let ignoreRegularTradingHours(ByVal Value As Boolean)
End Property

Private Property Get ignoreRegularTradingHours() As Boolean
End Property

Private Property Let limitPrice(ByVal Value As Double)
End Property

Private Property Get limitPrice() As Double
End Property

Private Property Let minimumQuantity(ByVal Value As Long)
End Property

Private Property Get minimumQuantity() As Long
End Property

Private Property Let NbboPriceCap(ByVal Value As Double)
End Property

Private Property Get NbboPriceCap() As Double
End Property

Private Property Let orderType(ByVal Value As OrderTypes)
End Property

Private Property Get orderType() As OrderTypes
End Property

Private Property Let origin(ByVal Value As OrderOrigins)
End Property

Private Property Get origin() As OrderOrigins
End Property

Private Property Let originatorRef(ByVal Value As String)
End Property

Private Property Get originatorRef() As String
End Property

Private Property Let overrideConstraints(ByVal Value As Boolean)
End Property

Private Property Get overrideConstraints() As Boolean
End Property

Private Property Let providerProperties(ByVal Value As TWUtilities30.Parameters)
End Property

Private Property Get providerProperties() As TWUtilities30.Parameters
End Property

Private Property Let quantity(ByVal Value As Long)
End Property

Private Property Get quantity() As Long
End Property

Private Property Let settlingFirm(ByVal Value As String)
End Property

Private Property Get settlingFirm() As String
End Property

Private Property Let status(ByVal Value As OrderStatuses)
End Property

Private Property Get status() As OrderStatuses
End Property

Private Property Let stopTriggerMethod(ByVal Value As StopTriggerMethods)
End Property

Private Property Get stopTriggerMethod() As StopTriggerMethods
End Property

Private Property Let sweepToFill(ByVal Value As Boolean)
End Property

Private Property Get sweepToFill() As Boolean
End Property

Private Property Let timeInForce(ByVal Value As OrderTifs)
End Property

Private Property Get timeInForce() As OrderTifs
End Property

Private Property Let tradeBuildId(ByVal Value As String)
End Property

Private Property Get tradeBuildId() As String
End Property

Private Property Let triggerPrice(ByVal Value As Double)
End Property

Private Property Get triggerPrice() As Double
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




