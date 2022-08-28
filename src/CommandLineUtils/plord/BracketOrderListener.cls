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

Implements IDeferredAction

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

Private WithEvents mBracketOrderListener            As OrderUtils27.BracketOrderListener
Attribute mBracketOrderListener.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mBracketOrderListener = New OrderUtils27.BracketOrderListener
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

mBracketOrderListener.Remove Data

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mBracketOrderListener Event Handlers
'@================================================================================

Private Sub mBracketOrderListener_BracketOrderAdded(ev As StateChangeEventData)
Const ProcName As String = "mBracketOrderListener_BracketOrderAdded"
On Error GoTo Err

Dim lBO As IBracketOrder: Set lBO = ev.Source
gWriteLineToConsole lBO.Key & _
                    ": bracket order state is: " & _
                    BracketOrderStateToString(lBO.State) & _
                    generateRolloverClause(lBO), _
                    True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBracketOrderListener_NotifyBracketOrderError(ev As BracketOrderErrorEventData)
Const ProcName As String = "mBracketOrderListener_NotifyBracketOrderError"
On Error GoTo Err

Dim lOrder As IOrder: Set lOrder = ev.AffectedOrder
gWriteLineToConsole generateMessage(lOrder, lOrder.ErrorMessage), True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBracketOrderListener_NotifyBracketOrderMessage(ev As BracketOrderMessageEventData)
Const ProcName As String = "mBracketOrderListener_NotifyBracketOrderMessage"
On Error GoTo Err

Dim lOrder As IOrder: Set lOrder = ev.AffectedOrder
gWriteLineToConsole generateMessage(lOrder, lOrder.Message), True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBracketOrderListener_Change(ev As ChangeEventData)
Const ProcName As String = "mBracketOrderListener_Change"
On Error GoTo Err

Dim lChangeType As BracketOrderChangeTypes: lChangeType = ev.ChangeType
Dim lBO As BracketOrder: Set lBO = ev.Source

Dim lRolloverClause As String: lRolloverClause = generateRolloverClause(lBO)

Dim s As String: s = lBO.Key & ": "
Dim s1 As String

Select Case lChangeType
Case BracketOrderCompleted
    s1 = "bracket order completed"
    If Not lBO.HasRolloverSpecification Then
        ' defer removing from the listener for a while to
        ' allow time for any late error messages to arrive -
        ' for example if the order size exceeds available
        ' funds
        DeferAction Me, lBO, 2, ExpiryTimeUnitSeconds
    End If
Case BracketOrderSelfCancelled
    s1 = "bracket order cancelled"
    mBracketOrderListener.Remove lBO
Case BracketOrderEntryOrderChanged

Case BracketOrderEntryOrderFilled
    s1 = "entry order filled " & _
        filledString(lBO.EntryOrder, lBO.Contract)
Case BracketOrderStopLossOrderChanged

Case BracketOrderStopLossOrderFilled
    s1 = "stop-loss order filled " & _
        filledString(lBO.StopLossOrder, lBO.Contract)
Case BracketOrderTargetOrderChanged

Case BracketOrderTargetOrderFilled
    s1 = "target order filled " & _
        filledString(lBO.TargetOrder, lBO.Contract)
Case BracketOrderCloseoutOrderCreated
    s = s & "closeout order created"
Case BracketOrderCloseoutOrderChanged

Case BracketOrderCloseoutOrderFilled
    s1 = "closeout order filled " & _
        filledString(lBO.CloseoutOrder, lBO.Contract)
Case BracketOrderSizeChanged

Case BracketOrderStateChanged
    s1 = "bracket order state changed: " & _
        BracketOrderStateToString(lBO.State) & _
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
        gGetContractName(lBO.NextContract)
Case BracketOrderRolloverCompleted
    s1 = "rollover completed: new bracket order is: " & _
        lBO.NextBracketOrder.Key
    Add lBO.NextBracketOrder
    mBracketOrderListener.Remove lBO
End Select

If s1 <> "" Then gWriteLineToConsole s & s1, True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Base() As OrderUtils27.BracketOrderListener
Set Base = mBracketOrderListener
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pBO As IBracketOrder)
Const ProcName As String = "Add"
On Error GoTo Err

mBracketOrderListener.Add pBO

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function filledString( _
                ByVal pOrder As IOrder, _
                ByVal pcontract As IContract) As String
filledString = pOrder.QuantityFilled & _
                " at " & FormatPrice(pOrder.LastFillPrice, _
                                    pcontract.Specifier.SecType, _
                                    pcontract.TickSize) & _
                "; remaining " & _
                pOrder.QuantityRemaining
End Function

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
    ": " & pMessage

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

