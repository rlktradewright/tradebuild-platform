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
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ChangeListener
Implements Task

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

Private Const ModuleName                            As String = "ClosePositionsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private WithEvents mOrderContexts                   As OrderContexts
Attribute mOrderContexts.VB_VarHelpID = -1

Private WithEvents mBalancingOrderContext           As OrderContext
Private WithEvents mBalancingBracketOrder           As BracketOrder
Attribute mBalancingBracketOrder.VB_VarHelpID = -1

Private mFlags                                      As ClosePositionsFlags

Private mOrderContextsWasEnabled                    As Boolean

Private mPositionManager                            As PositionManager

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
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_Run()
Const ProcName As String = "Task_Run"
On Error GoTo Err

Dim cancellingOrders As Boolean

If Not mBalancingOrderContext.IsReady Then
    mTaskContext.Suspend -1
    Exit Sub
End If

If (mFlags And ClosePositionsCancelOrders) = ClosePositionsCancelOrders Then
    ' first need to tell each OrderContext to cancel any outstanding stuff
    Dim lOrderContext As OrderContext
    For Each lOrderContext In mOrderContexts
        If lOrderContext.Cancel And (mFlags And ClosePositionsWaitForCancel) Then
            cancellingOrders = True
        End If
    Next
End If

mOrderContextsWasEnabled = mOrderContexts.Enabled
mOrderContexts.Enabled = False  ' prevent any further orders being placed
                                ' while the position is being closed

If Not cancellingOrders And mPositionManager.PositionSize <> 0 Then placeBalancingOrder

mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mBalancingBracketOrder Event Handlers
'@================================================================================

Private Sub mBalancingBracketOrder_Completed()
' now need to adjust the other BracketOrders to take account of the
' settling order

Const ProcName As String = "mBalancingBracketOrder_Completed"
On Error GoTo Err

'Dim oc As OrderContext
'
'For Each oc In mOrderContexts
'    oc.NotifyPositionCloseout mBalancingBracketOrder.EntryPrice, _
'                        mBalancingBracketOrder.StartTime
'Next

mOrderContexts.Enabled = mOrderContextsWasEnabled
mTaskContext.Finish Empty, False
 
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

If ev.changeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub

If mBalancingOrderContext.IsReady Then mTaskContext.Continue

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mOrderContexts Event Handlers
'@================================================================================

Private Sub mOrderContexts_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mOrderContexts_CollectionChanged"
On Error GoTo Err

Dim lOrderContext As OrderContext

Select Case ev.changeType
Case CollItemAdded
    Set lOrderContext = ev.AffectedItem
    lOrderContext.AddChangeListener Me
Case CollItemRemoved
    Set lOrderContext = ev.AffectedItem
    lOrderContext.RemoveChangeListener Me
End Select

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

Friend Sub Initialise( _
                ByVal pPositionManager As PositionManager, _
                ByVal pFlags As ClosePositionsFlags)
Set mPositionManager = pPositionManager
Set mOrderContexts = mPositionManager.OrderContexts
mFlags = pFlags

Set mBalancingOrderContext = mOrderContexts.Add("$balancing")
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub handleInactiveOrderContext(ByVal pOrderContext As OrderContext)
Const ProcName As String = "handleInactiveOrderContext"
On Error GoTo Err

If pOrderContext Is mBalancingOrderContext Then
    'should now be net flat
    mOrderContexts.Enabled = True
    
    If mPositionManager.EndingSession Then
        gLogMoneyManagement "Ending Session", mPositionManager.IsSimulated, Me
    End If
    
    mOrderContexts.Enabled = mOrderContextsWasEnabled
    mTaskContext.Finish Empty, False

Else
    Dim Size As Long
    
    Dim OrderContext As OrderContext
    For Each OrderContext In mOrderContexts
        If OrderContext.IsActive Then Exit Sub
        Size = Size + OrderContext.Size
    Next
    
    ' If we get here, all OrderContexts must be complete.
    
    ' Now, if there is a net position, need to close it out
    
    If Size = 0 Then
        If mPositionManager.EndingSession Then
            gLogMoneyManagement "Ending Session", mPositionManager.IsSimulated, Me
        End If
        mOrderContexts.Enabled = mOrderContextsWasEnabled
        mTaskContext.Finish Empty, False
    
    Else
        placeBalancingOrder
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub placeBalancingOrder()
Const ProcName As String = "placeBalancingOrder"
On Error GoTo Err

Dim lContract As IContract
Set lContract = mPositionManager.DataSource.ContractFuture.Value

gLogOrder _
        "Close positions at market: (Curr price=" & _
            FormatPrice(mPositionManager.DataSource.CurrentTick(TickTypeTrade).Price, _
                        lContract.Specifier.SecType, _
                        lContract.TickSize) & ")", _
        mPositionManager.IsSimulated, _
        Me

mBalancingOrderContext.Enabled = True
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




