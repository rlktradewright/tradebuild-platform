VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IBracketOrderErrorListener
Implements IBracketOrderMsgListener
Implements IChangeListener

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

Private Const ModuleName                            As String = "BracketOrderListener"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IBracketOrderErrorListener Interface Members
'@================================================================================

Private Sub IBracketOrderErrorListener_NotifyBracketOrderError(ev As BracketOrderErrorEventData)
Const ProcName As String = "IBracketOrderErrorListener_NotifyBracketOrderError"
On Error GoTo Err

Dim lOrder As IOrder: Set lOrder = ev.AffectedOrder
gWriteLineToConsole generateMessage(lOrder, lOrder.ErrorMessage), True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IBracketOrderMsgListener Interface Members
'@================================================================================

Private Sub IBracketOrderMsgListener_NotifyBracketOrderMessage(ev As BracketOrderMessageEventData)
Const ProcName As String = "IBracketOrderMsgListener_NotifyBracketOrderMessage"
On Error GoTo Err

Dim lOrder As IOrder: Set lOrder = ev.AffectedOrder
gWriteLineToConsole generateMessage(lOrder, lOrder.Message), True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

Dim lTime As String: lTime = FormatTimestamp(GetTimestamp, TimestampTimeOnlyISO8601) & " "

Dim lChangeType As BracketOrderChangeTypes: lChangeType = ev.ChangeType
Dim lBo As IBracketOrder: Set lBo = ev.Source
Select Case lChangeType
Case BracketOrderCompleted
    gWriteLineToStdOut lTime & lBo.Key & ": bracket order completed"
    Remove lBo
Case BracketOrderSelfCancelled
    gWriteLineToStdOut lTime & lBo.Key & ": bracket order cancelled"
    Remove lBo
Case BracketOrderEntryOrderChanged

Case BracketOrderEntryOrderFilled
    gWriteLineToStdOut lTime & lBo.Key & ": entry order filled " & lBo.EntryOrder.QuantityFilled & "; remaining " & lBo.EntryOrder.QuantityRemaining
Case BracketOrderStopLossOrderChanged

Case BracketOrderStopLossOrderFilled
    gWriteLineToStdOut lTime & lBo.Key & ": stop-loss order filled " & lBo.StopLossOrder.QuantityFilled & "; remaining " & lBo.StopLossOrder.QuantityRemaining
Case BracketOrderTargetOrderChanged

Case BracketOrderTargetOrderFilled
    gWriteLineToStdOut lTime & lBo.Key & ": target order filled " & lBo.TargetOrder.QuantityFilled & "; remaining " & lBo.TargetOrder.QuantityRemaining
Case BracketOrderCloseoutOrderCreated
    gWriteLineToStdOut lTime & lBo.Key & ": closeout order created"
Case BracketOrderCloseoutOrderChanged

Case BracketOrderCloseoutOrderFilled
    gWriteLineToStdOut lTime & lBo.Key & ": closeout order filled " & lBo.CloseoutOrder.QuantityFilled & "; remaining " & lBo.CloseoutOrder.QuantityRemaining
Case BracketOrderSizeChanged

Case BracketOrderStateChanged
    gWriteLineToStdOut lTime & lBo.Key & ": bracket order state changed: " & BracketOrderStateToString(lBo.State)
Case BracketOrderChangesPending

Case BracketOrderChangesApplied

Case BracketOrderChangesCancelled

Case BracketOrderPositionCloseoutNotified

End Select

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pBO As IBracketOrder)
pBO.AddBracketOrderErrorListener Me
pBO.AddBracketOrderMessageListener Me
pBO.AddChangeListener Me
End Sub

Friend Sub Remove(ByVal pBO As IBracketOrder)
pBO.RemoveBracketOrderErrorListener Me
pBO.RemoveBracketOrderMessageListener Me
pBO.RemoveChangeListener Me
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function generateMessage( _
                ByVal pOrder As IOrder, _
                ByVal pMessage As String) As String
Const ProcName As String = "generateMessage"
On Error GoTo Err

Dim s As String
s = pOrder.GroupName & ": " & _
    "(" & pOrder.ContractSpecifier.LocalSymbol & "@" & pOrder.ContractSpecifier.Exchange & ")" & _
    " " & OrderActionToString(pOrder.Action) & _
    " " & pOrder.Quantity & _
    " " & OrderTypeToShortString(pOrder.OrderType) & _
    ": " & pOrder.ErrorMessage

generateMessage = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


