VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClosePositionsTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements ChangeListener

'@================================================================================
' Events
'@================================================================================

Event PositionClosed()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ClosePositionsTask"

Private Const BalancingOrderContextName             As String = "$balancing"
'@================================================================================
' Member variables
'@================================================================================

Private mOrderContexts                              As OrderContexts
Attribute mOrderContexts.VB_VarHelpID = -1

Private WithEvents mBalancingOrderContext           As OrderContext
Attribute mBalancingOrderContext.VB_VarHelpID = -1
Private WithEvents mBalancingBracketOrder           As BracketOrder
Attribute mBalancingBracketOrder.VB_VarHelpID = -1

Private mFlags                                      As ClosePositionsFlags

Private mOrderContextsWasEnabled                    As Boolean

Private mPositionManager                            As PositionManager

Private mLastExecutionReport                        As IExecutionReport

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "ChangeListener_Change"
On Error GoTo Err

Dim lOrderContext As OrderContext
Set lOrderContext = ev.Source
If ev.changeType = OrderContextChangeTypes.OrderContextActiveStateChanged And _
    (Not lOrderContext.IsActive) _
Then
    handleInactiveOrderContext lOrderContext
    lOrderContext.RemoveChangeListener Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mBalancingBracketOrder Event Handlers
'@================================================================================

Private Sub mBalancingBracketOrder_Completed()
Const ProcName As String = "mBalancingBracketOrder_Completed"
On Error GoTo Err

' adjust the other BracketOrders to take account of the settling Order

Dim oc As OrderContext
For Each oc In mOrderContexts
    If oc.Name <> BalancingOrderContextName Then oc.NotifyPositionCloseout mLastExecutionReport
Next

Assert mPositionManager.PositionSize = 0, "Position size is not zero"

doFinalActions
 
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBalancingBracketOrder_Execution(ev As ExecutionEventData)
Const ProcName As String = "mBalancingBracketOrder_Execution"
On Error GoTo Err

Set mLastExecutionReport = ev.ExecutionReport

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mBalancingOrderContext Event Handlers
'@================================================================================

Private Sub mBalancingOrderContext_Change(ev As ChangeEventData)
Const ProcName As String = "mBalancingOrderContext_Change"
On Error GoTo Err

If ev.changeType = OrderContextChangeTypes.OrderContextReadyStateChanged And _
    mBalancingOrderContext.IsReady _
Then
    placeBalancingOrder
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub ClosePositions( _
                ByVal pPositionManager As PositionManager, _
                ByVal pFlags As ClosePositionsFlags)
Const ProcName As String = "ClosePositions"
On Error GoTo Err

Set mPositionManager = pPositionManager
Set mOrderContexts = mPositionManager.OrderContexts
mFlags = pFlags

mOrderContextsWasEnabled = mOrderContexts.Enabled
mOrderContexts.Enabled = False  ' prevent any further orders being placed
                                ' while the position is being closed

If Not (mFlags And ClosePositionsCancelOrders) = ClosePositionsCancelOrders Then
    placeBalancingOrder
    Exit Sub
End If

Dim lCancelling As Boolean
lCancelling = cancelOrders

If Not lCancelling Or Not (mFlags And ClosePositionsWaitForCancel) = ClosePositionsWaitForCancel Then
    ' either nothing to cancel or we're not waiting
    ' for cancellations to complete
    placeBalancingOrder
    Exit Sub
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function cancelOrders() As Boolean
Const ProcName As String = "cancelOrders"
On Error GoTo Err

Dim lOrderContext As OrderContext
For Each lOrderContext In mOrderContexts
    If lOrderContext.Cancel Then
        cancelOrders = True
        If (mFlags And ClosePositionsWaitForCancel) = ClosePositionsWaitForCancel Then lOrderContext.AddChangeListener Me
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub doFinalActions()
mOrderContexts.Enabled = mOrderContextsWasEnabled
If mPositionManager.EndingSession Then
    gLogMoneyManagement "Ending Session", mPositionManager.IsSimulated, Me
End If
RaiseEvent PositionClosed
End Sub

Private Sub handleInactiveOrderContext(ByVal pOrderContext As OrderContext)
Const ProcName As String = "handleInactiveOrderContext"
On Error GoTo Err

Dim Size As Long

Dim OrderContext As OrderContext
For Each OrderContext In mOrderContexts
    If OrderContext.IsActive Then Exit Sub
    Size = Size + OrderContext.Size
Next

' If we get here, all OrderContexts must be complete.

' Now, if there is a net position, need to close it out

If Size = 0 Then
    doFinalActions
Else
    placeBalancingOrder
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub placeBalancingOrder()
Const ProcName As String = "placeBalancingOrder"
On Error GoTo Err

If mPositionManager.PositionSize = 0 Then
    doFinalActions
    Exit Sub
End If

If mBalancingOrderContext Is Nothing Then
    ' use a local variable here to avoid catching the ReadyStateChanged event
    Dim lBalancingOrderContext As OrderContext
    Set lBalancingOrderContext = mOrderContexts.Add(BalancingOrderContextName)
    lBalancingOrderContext.Enabled = True
    
    Set mBalancingOrderContext = lBalancingOrderContext
End If
If Not mBalancingOrderContext.IsReady Then Exit Sub

Dim lContract As IContract
Set lContract = mPositionManager.DataSource.ContractFuture.Value

gLogOrder _
        "Close positions at market: (Curr price=" & _
            FormatPrice(mPositionManager.DataSource.CurrentTick(TickTypeTrade).Price, _
                        lContract.Specifier.SecType, _
                        lContract.TickSize) & ")", _
        mPositionManager.IsSimulated, _
        Me

If mBalancingOrderContext.IsOrderTifSupported(OrderTIFs.OrderTIFDay) Then mBalancingOrderContext.TemplateOrder.TimeInForce = OrderTIFs.OrderTIFDay
Set mBalancingBracketOrder = _
                mBalancingOrderContext.CreateBracketOrder( _
                            IIf(mPositionManager.PositionSize < 0, OrderActionBuy, OrderActionSell), _
                            Abs(mPositionManager.PositionSize), _
                            BracketEntryTypes.BracketEntryTypeMarket, _
                            0#, _
                            0, _
                            0#, _
                            BracketStopLossTypes.BracketStopLossTypeNone, _
                            0#, _
                            0, _
                            0#, _
                            BracketTargetTypes.BracketTargetTypeNone, _
                            0#, _
                            0, _
                            0#)
mBalancingOrderContext.ExecuteBracketOrder mBalancingBracketOrder

Set mBalancingOrderContext = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




