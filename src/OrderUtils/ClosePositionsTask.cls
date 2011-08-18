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

Private mBalancingOrderContext                      As OrderContext
Private WithEvents mBalancingBracketOrder           As BracketOrder
Attribute mBalancingBracketOrder.VB_VarHelpID = -1

Private mflags                                      As ClosePositionsFlags

Private mOrderContextsWasEnabled                    As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEventData)
Dim lOrderContext As OrderContext
Set lOrderContext = ev.Source
If ev.changeType = OrderContextChangeTypes.OrderContextActiveStateChanged And _
    (Not lOrderContext.IsActive) _
Then
    handleInactiveOrderContext lOrderContext
End If
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_Run()
Dim lOrderContext As OrderContext
Dim cancellingOrders As Boolean

Const ProcName As String = "Task_Run"
On Error GoTo Err

If (mflags And ClosePositionsCancelOrders) = ClosePositionsCancelOrders Then
    ' first need to tell each OrderContext to cancel any outstanding stuff
    For Each lOrderContext In mOrderContexts
        If lOrderContext.Cancel And (mflags And ClosePositionsWaitForCancel) Then
            cancellingOrders = True
        End If
    Next
End If

mOrderContextsWasEnabled = mOrderContexts.Enabled
mOrderContexts.Enabled = False  ' prevent any further orders being placed
                                ' while the position is being closed

If Not cancellingOrders And PositionSize <> 0 Then placeBalancingOrder

mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Dim oc As OrderContext

For Each oc In mOrderContexts
    oc.NotifyPositionCloseout mBalancingBracketOrder.EntryPrice, _
                        mBalancingBracketOrder.StartTime
Next

mOrderContexts.Enabled = mOrderContextsWasEnabled
mTaskContext.Finish Empty, False
 
Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mOrderContexts Event Handlers
'@================================================================================

Private Sub mOrderContexts_CollectionChanged(ev As CollectionChangeEventData)
Dim lOrderContext As OrderContext
Select Case ev.changeType
Case CollItemAdded
    Set lOrderContext = ev.AffectedItem
    lOrderContext.AddChangeListener Me
Case CollItemRemoved
    Set lOrderContext = ev.AffectedItem
    lOrderContext.RemoveChangeListener Me
End Select
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pOrderContexts As OrderContexts, _
                ByVal pFlags As ClosePositionsFlags)
Set mOrderContexts = pOrderContexts
mflags = pFlags

Set mBalancingOrderContext = mOrderContexts.Add("$balancing")
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub handleInactiveOrderContext(ByVal pOrderContext As OrderContext)
Const ProcName As String = "handleInactiveOrderContext"
On Error GoTo Err

Dim OrderContext As OrderContext
Dim size As Long

If pOrderContext Is mBalancingOrderContext Then
    'should now be net flat
    mOrderContexts.Enabled = True
    
    If mEndSession Then
        GLogMoneyManagement "Ending Session", mIsSimulated, Me
    End If
    
    mOrderContexts.Enabled = mOrderContextsWasEnabled
    mTaskContext.Finish Empty, False

Else
    
    For Each OrderContext In mOrderContexts
        If OrderContext.IsActive Then Exit Sub
        size = size + OrderContext.size
    Next
    
    ' If we get here, all OrderContexts must be complete.
    
    ' Now, if there is a net position, need to close it out
    
    If size = 0 Then
        If mEndSession Then
            GLogMoneyManagement "Ending Session", mIsSimulated, Me
        End If
        mOrderContexts.Enabled = mOrderContextsWasEnabled
        mTaskContext.Finish Empty, False
    
    Else
        placeBalancingOrder
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub placeBalancingOrder()
Const ProcName As String = "placeBalancingOrder"

On Error GoTo Err

GLogOrder _
        "Close positions at market: (Curr price=" & mDataSource.FormatPrice(mDataSource.TradePrice) & ")", _
        mIsSimulated, _
        Me

mBalancingOrderContext.Enabled = True
If mBalancingOrderContext.IsOrderTifValueSupported(TIFDay) Then mBalancingOrderContext.TimeInForce = TIFDay
Set mBalancingBracketOrder = mBalancingOrderContext.CreateBracketOrder( _
                            IIf(PositionSize < 0, ActionBuy, ActionSell), _
                            Abs(PositionSize), _
                            BracketEntryTypes.EntryOrderTypeMarket, _
                            0, _
                            0, _
                            0, _
                            StopOrderTypes.StopOrderTypeNone, _
                            0, _
                            0, _
                            0, _
                            BracketTargetTypes.TargetOrderTypeNone, _
                            0, _
                            0, _
                            0)
mBalancingOrderContext.ExecuteBracketOrder mBalancingBracketOrder

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub




