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

Dim lOrder As Iorder: Set lOrder = ev.AffectedOrder
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

Dim lOrder As Iorder: Set lOrder = ev.AffectedOrder
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
Dim lBo As BracketOrder: Set lBo = ev.Source

Dim lRolloverClause As String: lRolloverClause = generateRolloverClause(lBo)

Dim s As String: s = lTime & lBo.Key & ": "
Dim s1 As String

Select Case lChangeType
Case BracketOrderCompleted
    s1 = "bracket order completed"
    If Not lBo.HasRolloverSpecification Then Remove lBo
Case BracketOrderSelfCancelled
    s1 = "bracket order cancelled"
    Remove lBo
Case BracketOrderEntryOrderChanged

Case BracketOrderEntryOrderFilled
    s1 = "entry order filled " & _
        filledString(lBo.EntryOrder, lBo.Contract)
Case BracketOrderStopLossOrderChanged

Case BracketOrderStopLossOrderFilled
    s1 = "stop-loss order filled " & _
        filledString(lBo.StopLossOrder, lBo.Contract)
Case BracketOrderTargetOrderChanged

Case BracketOrderTargetOrderFilled
    s1 = "target order filled " & _
        filledString(lBo.TargetOrder, lBo.Contract)
Case BracketOrderCloseoutOrderCreated
    s = s & "closeout order created"
Case BracketOrderCloseoutOrderChanged

Case BracketOrderCloseoutOrderFilled
    s1 = "closeout order filled " & _
        filledString(lBo.CloseoutOrder, lBo.Contract)
Case BracketOrderSizeChanged

Case BracketOrderStateChanged
    s1 = "bracket order state changed: " & _
        BracketOrderStateToString(lBo.State) & _
        lRolloverClause
Case BracketOrderChangesPending

Case BracketOrderChangesApplied

Case BracketOrderChangesCancelled

Case BracketOrderPositionCloseoutNotified
    s1 = "closed out"
Case BracketOrderRolloverInitiated
    s1 = "bracket order rollover initiated"
Case BracketOrderRolloverGotContract
    s1 = "bracket order rolling over to next contract: " & _
        gGetContractName(lBo.NextContract)
Case BracketOrderRolloverCompleted
    s1 = "new bracket order is: " & _
        lBo.NextBracketOrder.Key
    Add lBo.NextBracketOrder
    Remove lBo
End Select

If s1 <> "" Then gWriteLineToStdOut s & s1

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

Dim lTime As String: lTime = FormatTimestamp(GetTimestamp, TimestampTimeOnlyISO8601) & " "
gWriteLineToStdOut lTime & _
                    pBO.Key & _
                    ": bracket order state is: " & _
                    BracketOrderStateToString(pBO.State) & _
                    generateRolloverClause(pBO)
End Sub

Friend Sub Remove(ByVal pBO As IBracketOrder)
pBO.RemoveBracketOrderErrorListener Me
pBO.RemoveBracketOrderMessageListener Me
pBO.RemoveChangeListener Me
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function filledString( _
                ByVal pOrder As Iorder, _
                ByVal pcontract As IContract) As String
filledString = pOrder.QuantityFilled & _
                " at " & FormatPrice(pOrder.LastFillPrice, _
                                    pcontract.Specifier.SecType, _
                                    pcontract.TickSize) & _
                "; remaining " & _
                pOrder.QuantityRemaining
End Function

Private Function generateMessage( _
                ByVal pOrder As Iorder, _
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

Private Function generateRolloverClause(ByVal pBO As BracketOrder) As String
Dim s As String
If pBO.HasRolloverSpecification Then
    If pBO.IsRolloverInProgress Then
        s = ": rollover in progress"
    Else
        s = ": rollover at " & _
            FormatTimestamp(pBO.RolloverDate, _
                TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
    End If
End If
generateRolloverClause = s
End Function

