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
' IDeferredAction Interface Members
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
LogMessage lBO.Key & ": " & _
                    gGetContractName(lBO.Contract) & _
                    ": state=" & _
                    BracketOrderStateToString(lBO.State) & _
                    generateRolloverClause(lBO)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBracketOrderListener_NotifyBracketOrderError(ev As BracketOrderErrorEventData)
Const ProcName As String = "mBracketOrderListener_NotifyBracketOrderError"
On Error GoTo Err

Dim lBO As IBracketOrder: Set lBO = ev.Source
Dim lOrder As IOrder: Set lOrder = ev.AffectedOrder
gWriteLineToConsole generateMessage(lBO, lOrder, lOrder.ErrorMessage, True), True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBracketOrderListener_NotifyBracketOrderMessage(ev As BracketOrderMessageEventData)
Const ProcName As String = "mBracketOrderListener_NotifyBracketOrderMessage"
On Error GoTo Err

Dim lBO As IBracketOrder: Set lBO = ev.Source
Dim lOrder As IOrder: Set lOrder = ev.AffectedOrder
gWriteLineToConsole generateMessage(lBO, lOrder, lOrder.Message, False), True

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

Dim s As String: s = lBO.Key & ": " & _
                    gGetContractName(lBO.Contract) & ": "

Dim s1 As String
Dim lToConsole As Boolean

Select Case lChangeType
Case BracketOrderCompleted
    lToConsole = True
    s1 = "bracket order completed"
    If Not lBO.HasRolloverSpecification Then
        ' defer removing from the listener for a while to
        ' allow time for any late error messages to arrive -
        ' for example if the order size exceeds available
        ' funds
        DeferAction Me, lBO, 2, ExpiryTimeUnitSeconds
    End If
Case BracketOrderSelfCancelled
    lToConsole = True
    s1 = "bracket order cancelled"
    mBracketOrderListener.Remove lBO
Case BracketOrderEntryOrderChanged

Case BracketOrderEntryOrderFilled
    lToConsole = True
    s1 = "entry order filled " & _
        filledString(lBO.EntryOrder, lBO.Contract)
Case BracketOrderStopLossOrderChanged

Case BracketOrderStopLossOrderFilled
    lToConsole = True
    s1 = "stop-loss order filled " & _
        filledString(lBO.StopLossOrder, lBO.Contract)
Case BracketOrderTargetOrderChanged

Case BracketOrderTargetOrderFilled
    lToConsole = True
    s1 = "target order filled " & _
        filledString(lBO.TargetOrder, lBO.Contract)
Case BracketOrderCloseoutOrderCreated
    lToConsole = True
    s = s & "closeout order created"
Case BracketOrderCloseoutOrderChanged

Case BracketOrderCloseoutOrderFilled
    lToConsole = True
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
    lToConsole = True
    s1 = "closed out"
Case BracketOrderChangeTypes.BracketOrderRolloverPending
    lToConsole = True
    s1 = "bracket order rollover pending" & _
        lRolloverClause
Case BracketOrderRolloverInitiated
    lToConsole = True
    s1 = "bracket order rollover initiated"
Case BracketOrderRolloverGotContract
    lToConsole = True
    s1 = "bracket order rolling over to next contract: " & _
        gGetContractName(lBO.NextContract)
Case BracketOrderRolloverCompleted
    lToConsole = True
    Dim lNewBo As IBracketOrder: Set lNewBo = lBO.NextBracketOrder
    s1 = "rollover completed: new bracket order is: " & _
        lNewBo.Key
    Add lNewBo
    CreateBracketProfitCalculator lNewBo, lNewBo.OrderContext.DataSource
    mBracketOrderListener.Remove lBO
End Select

If s1 <> "" Then
    If lToConsole Then
        gWriteLineToConsole s & s1, True
    Else
        LogMessage s & s1
    End If
End If
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

Friend Sub Add(ByVal pBo As IBracketOrder)
Const ProcName As String = "Add"
On Error GoTo Err

mBracketOrderListener.Add pBo

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
                ByVal pBo As IBracketOrder, _
                ByVal pOrder As IOrder, _
                ByVal pMessage As String, _
                ByVal pIsError As Boolean) As String
Const ProcName As String = "generateMessage"
On Error GoTo Err

Dim s As String
s = pBo.Key & _
    ": " & OrderActionToString(pOrder.Action) & _
    " " & pOrder.Quantity & _
    " " & OrderTypeToShortString(pOrder.OrderType) & _
    IIf(pIsError, ": error", "") & _
    ": " & pMessage

generateMessage = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateRolloverClause(ByVal pBo As BracketOrder) As String
Dim s As String
If Not pBo.HasRolloverSpecification Then
ElseIf pBo.IsRolloverInProgress Then
    s = ": rollover in progress"
ElseIf pBo.State = BracketOrderStateClosed And pBo.Size = 0 Then
Else
    s = ": rollover at " & _
        FormatTimestamp(pBo.RolloverDate, _
            TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
End If
generateRolloverClause = s
End Function

