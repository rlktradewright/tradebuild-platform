VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBracketOrder"
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

Private Const ModuleName                            As String = "IBracketOrder"

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

Public Property Get ApplicationIndex() As Long
End Property

Public Property Let CancelAfter(ByVal Value As Long)
End Property

Public Property Let CancelPrice(ByVal Value As Double)
End Property

Public Property Get CancelPrice() As Double
End Property

Public Property Let CancelTime(ByVal Value As Date)
End Property

Public Property Get CancelTime() As Date
End Property

Public Property Get CloseoutOrder() As IOrder
End Property

Public Property Get ContextName() As String
End Property

Public Property Get Contract() As IContract
End Property

Public Property Get CreationTime() As Date
End Property

Public Property Get CumBuyPrice() As Double
End Property

Public Property Get CumSellPrice() As Double
End Property

Public Property Let Description(ByVal Value As String)
End Property

Public Property Get Description() As String
End Property

Public Property Get EndTime() As Date
End Property

Public Property Get EntryOrder() As IOrder
End Property

Public Property Get EntryPrice() As Double
End Property

Public Property Let EntryReason(ByVal Value As String)
End Property

Public Property Get EntryReason() As String
End Property

Public Property Get ExitPrice() As Double
End Property

Public Property Get GroupName() As String
End Property

Public Property Get IsComplete() As Boolean
End Property

Public Property Get IsDirty() As Boolean
End Property

Public Property Get IsSimulated() As Boolean
End Property

Public Property Get Key() As String
End Property

Public Property Get LongPosition() As Boolean
End Property

Public Property Get Order(ByVal pRole As BracketOrderRoles) As IOrder
End Property

Public Property Get OrderContext() As OrderContext
End Property

Public Property Get PendingSize() As BoxedDecimal
End Property

Public Property Let PreventUnprotectedPosition(ByVal Value As Boolean)
End Property

Public Property Get PreventUnprotectedPosition() As Boolean
End Property

Public Property Get Risk() As Currency
End Property

Public Property Let RolloverSpecification(ByVal Value As RolloverSpecification)
End Property

Public Property Get RolloverSpecification() As RolloverSpecification
End Property

Public Property Get Size() As BoxedDecimal
End Property

Public Property Get StartTime() As Date
End Property

Public Property Get State() As BracketOrderStates
End Property

Public Property Get StopLossOrder() As IOrder
End Property

Public Property Let StopReason(ByVal Value As String)
End Property

Public Property Get StopReason() As String
End Property

Public Property Get TargetOrder() As IOrder
End Property

Public Property Let TargetReason(ByVal Value As String)
End Property

Public Property Get TargetReason() As String
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddBracketOrderErrorListener(ByVal pListener As IBracketOrderErrorListener)
End Sub

Public Sub AddBracketOrderMessageListener(ByVal pListener As IBracketOrderMsgListener)
End Sub

Public Sub AddChangeListener(ByVal pListener As IChangeListener)
End Sub

Public Sub AddExecutionListener(ByVal pListener As IExecutionListener)
End Sub

Public Sub AddStateChangeListener(ByVal pListener As IStateChangeListener)
End Sub

Public Sub AdjustStop(ByVal pStopTriggerPriceSpec As PriceSpecifier, Optional ByVal pQuantity As BoxedDecimal)
End Sub

Public Sub AdjustTarget(ByVal pTargetPriceSpec As PriceSpecifier, Optional ByVal pQuantity As BoxedDecimal)
End Sub

Public Function Cancel(Optional ByVal pEvenIfFilled As Boolean = False) As Boolean
End Function

Public Sub CancelChanges()
End Sub

Public Sub Checkpoint()
End Sub

Public Sub CloseOut(Optional ByVal pCloseoutSpec As CloseoutSpecifier)
End Sub

Public Function ContainsOrder(ByVal pOrder As IOrder) As Boolean
End Function

Public Sub Execute()
End Sub

Public Sub NotifyCloseout(ByVal pExec As IExecutionReport)
End Sub

Public Sub RemoveBracketOrderErrorListener(ByVal pListener As IBracketOrderErrorListener)
End Sub

Public Sub RemoveBracketOrderMessageListener(ByVal pListener As IBracketOrderMsgListener)
End Sub

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)
End Sub

Public Sub RemoveExecutionListener(ByVal pListener As IExecutionListener)
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
End Sub

Public Sub SetNewEntryLimitPrice( _
                ByVal pPriceSpec As PriceSpecifier)
End Sub

Public Sub SetNewEntryTriggerPrice( _
                ByVal pPriceSpec As PriceSpecifier)
End Sub

Public Sub SetNewOrderLimitPrice( _
                ByVal pRole As BracketOrderRoles, _
                ByVal pPriceSpec As PriceSpecifier)
End Sub

Public Sub SetNewOrderTriggerPrice( _
                ByVal pRole As BracketOrderRoles, _
                ByVal pPriceSpec As PriceSpecifier)
End Sub

Public Sub SetNewEntryQuantity(ByVal pNewQuantity As BoxedDecimal)
End Sub

Public Sub SetNewStopLossLimitPrice( _
                ByVal pPriceSpec As PriceSpecifier)
End Sub

Public Sub SetNewStopLossQuantity(ByVal pNewQuantity As BoxedDecimal)
End Sub

Public Sub SetNewStopLossTriggerPrice( _
                ByVal pPriceSpec As PriceSpecifier)
End Sub

Public Sub SetNewTargetLimitPrice( _
                ByVal pPriceSpec As PriceSpecifier)
End Sub

Public Sub SetNewTargetQuantity(ByVal pNewQuantity As BoxedDecimal)
End Sub

Public Sub SetNewTargetTriggerPrice( _
                ByVal pPriceSpec As PriceSpecifier)
End Sub

Public Function ToString() As String
End Function

Public Sub Update()
End Sub

'@================================================================================
' Helper Functions
'@================================================================================






