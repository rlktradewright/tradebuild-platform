VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionCloser"
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

Implements IChangeListener
Implements IDeferredAction

'@================================================================================
' Events
'@================================================================================

Event PositionNotClosed(ByVal pErrorMessage As String)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "PositionCloser"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderContexts                              As OrderContexts
Attribute mOrderContexts.VB_VarHelpID = -1

Private WithEvents mBalancingOrderContext           As OrderContext
Attribute mBalancingOrderContext.VB_VarHelpID = -1
Private WithEvents mBalancingBracketOrder           As BracketOrder
Attribute mBalancingBracketOrder.VB_VarHelpID = -1

Private mOrderContextsWasEnabled                    As Boolean

Private mPositionManager                            As PositionManager
Private mContract                                   As IContract
Private mClock                                      As Clock

Private mLastExecutionReport                        As IExecutionReport

Private mUseLimitorder                              As Boolean
Private mBidAskSpreadPercent                        As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
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
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

If mPositionManager.HasBracketOrdersWithNonZeroSizes Then
ElseIf mPositionManager.HasBracketOrdersWithNonZeroPendingSizes Then
Else
    Exit Sub
End If

Set mOrderContexts = mPositionManager.OrderContexts
Set mContract = mPositionManager.ContractFuture.Value
Set mClock = mPositionManager.DataSource.ClockFuture.Value

mOrderContextsWasEnabled = mOrderContexts.Enabled
mOrderContexts.Enabled = False  ' prevent any further orders being placed
                                ' while the position is being closed

If Not cancelOrders Then ensureAllBracketOrdersClosedOut

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

mBalancingBracketOrder.SetSize 0

' adjust the other BracketOrders to take account of the settling Order

If mLastExecutionReport Is Nothing Then
    ' the balancing bracket order has failed due to some error
    doFinalActions
    RaiseEvent PositionNotClosed(mBalancingBracketOrder.EntryOrder.ErrorMessage)
    Exit Sub
End If
    
Dim oc As OrderContext
Dim i As Long
For Each oc In mOrderContexts
    If Not oc Is mBalancingOrderContext Then
        Dim lBracketOrder As IBracketOrder
        For Each lBracketOrder In oc.ClosedBracketOrders
            If lBracketOrder.Size <> 0 Then
                
                ' NB: we don't save recovery data for this execution
                ' because it will simply be recreated at recovery
                
                Dim lExec As Execution
                Set lExec = oc.NewExecution(True)
                lExec.GroupName = oc.GroupName
                lExec.RecoveryContextName = oc.RecoveryContextName
                lExec.SyncToExecutionReport mLastExecutionReport
                
                ' give this execution a unique id
                i = i + 1
                lExec.Id = lExec.Id & "{" & i & "}"
    
                lExec.Quantity = Abs(lBracketOrder.Size)
                If lBracketOrder.Size > 0 Then
                    lExec.Action = OrderActionSell
                Else
                    lExec.Action = OrderActionBuy
                End If
                
                lBracketOrder.NotifyCloseout lExec
            End If
        Next
    End If
Next

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

If ev.changeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub

If mBalancingOrderContext.IsReadyForBracketOrder( _
            IIf(mUseLimitorder, _
                BracketEntryTypes.BracketEntryTypeLimit, _
                BracketEntryTypes.BracketEntryTypeMarket)) Then
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
                ByVal pUseLimitOrder As Boolean, _
                ByVal pBidAskSpreadPercent As Long)
Const ProcName As String = "ClosePositions"
On Error GoTo Err

Set mPositionManager = pPositionManager
mUseLimitorder = pUseLimitOrder
mBidAskSpreadPercent = pBidAskSpreadPercent

DeferAction Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RecoverBalancingOrder( _
                ByVal pBracketOrder As BracketOrder, _
                ByVal pPositionManager As PositionManager)
Const ProcName As String = "RecoverBalancingOrder"
On Error GoTo Err

Set mBalancingBracketOrder = pBracketOrder
Set mPositionManager = pPositionManager
Set mOrderContexts = mPositionManager.OrderContexts

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
        lOrderContext.AddChangeListener Me
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub doFinalActions()
Const ProcName As String = "doFinalActions"
On Error GoTo Err

mOrderContexts.Enabled = mOrderContextsWasEnabled
If mPositionManager.EndingSession Then
    gLogMoneyManagement "Ending Session", mPositionManager.IsSimulated, Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ensureAllBracketOrdersClosedOut()
Const ProcName As String = "ensureAllBracketOrdersClosedOut"
On Error GoTo Err

If mPositionManager.PositionSize <> 0 Then
    placeBalancingOrder
Else
    ' we have some bracket orders that net out to no position
    ' but have not been closed out, so we need to mark them
    ' to market
    markToMarket
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub handleInactiveOrderContext(ByVal pOrderContext As OrderContext)
Const ProcName As String = "handleInactiveOrderContext"
On Error GoTo Err

If mOrderContexts.HasActiveBracketOrders Then Exit Sub

' If we get here, all bracket orders must be closed.

ensureAllBracketOrdersClosedOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub markToMarket()
Const ProcName As String = "markToMarket"
On Error GoTo Err

' closeout non-zero-size bracket orders marking them to mnarket

If Not mPositionManager.DataSource.HasCurrentTick(TickTypeBid) Or _
    Not mPositionManager.DataSource.HasCurrentTick(TickTypeAsk) Then
    RaiseEvent PositionNotClosed("No bid or ask tick available")
    Exit Sub
End If

Dim lBidPrice As Double: lBidPrice = mPositionManager.DataSource.CurrentTick(TickTypeBid).Price
Dim lAskPrice As Double: lBidPrice = mPositionManager.DataSource.CurrentTick(TickTypeAsk).Price

Dim oc As OrderContext
For Each oc In mOrderContexts
    If Not oc Is mBalancingOrderContext Then
        Dim lBracketOrder As IBracketOrder
        For Each lBracketOrder In oc.ClosedBracketOrders
            If lBracketOrder.Size <> 0 Then
                Dim lExec As Execution: Set lExec = oc.NewExecution(True)
                lExec.GroupName = oc.GroupName
                lExec.RecoveryContextName = oc.RecoveryContextName
                lExec.Quantity = Abs(lBracketOrder.Size)
                If lBracketOrder.Size > 0 Then
                    lExec.Action = OrderActionSell
                Else
                    lExec.Action = OrderActionBuy
                End If
                lExec.FillingExchange = "TRADEBUILD"
                lExec.FillTime = mClock.Timestamp
                lExec.Id = "MarkToMarket"
                lExec.Price = mPositionManager.DataSource.CurrentTick(TickTypeTrade).Price
                lExec.SecurityName = mContract.Specifier.LocalSymbol & "@" & mContract.Specifier.Exchange
                lExec.TimezoneName = mContract.TimezoneName
                
                lBracketOrder.NotifyCloseout lExec
            
                lExec.SaveRecoveryInfo
            End If
        Next
    End If
Next

doFinalActions

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub placeBalancingOrder()
Const ProcName As String = "placeBalancingOrder"
On Error GoTo Err

Set mBalancingOrderContext = mOrderContexts.BalancingOrderContext
If Not mBalancingOrderContext.IsReadyForBracketOrder( _
            IIf(mUseLimitorder, _
                BracketEntryTypes.BracketEntryTypeLimit, _
                BracketEntryTypes.BracketEntryTypeMarket)) Then
    Exit Sub
End If

Dim lContract As IContract
Set lContract = mPositionManager.DataSource.ContractFuture.Value

If mUseLimitorder Then
    Dim lBid As Double: lBid = mPositionManager.DataSource.CurrentTick(TickTypeBid).Price
    Dim lAsk As Double: lAsk = mPositionManager.DataSource.CurrentTick(TickTypeAsk).Price
    
    Dim lTicksize As Double
    If lContract.Specifier.SecType = SecTypeOption Or _
                        lContract.Specifier.SecType = SecTypeFuturesOption _
    Then
        lTicksize = GetTickSizeForOptions(lBid)
    ElseIf lContract.Specifier.SecType = SecTypeStock Then
        lTicksize = GetTickSizeForStocks(lContract.Specifier.Exchange, lBid)
    End If
    If lTicksize = 0# Then lTicksize = lContract.TickSize
    
    Dim lPriceOffset As Double: lPriceOffset = Int(mBidAskSpreadPercent * (CLng(lAsk / lTicksize) - CLng(lBid / lTicksize)) / 100#) * lTicksize
    Dim lEntryPrice As Double
    If mPositionManager.PositionSize < 0 Then
        lEntryPrice = lBid + lPriceOffset
    Else
        lEntryPrice = lAsk - lPriceOffset
    End If
    Set mBalancingBracketOrder = _
                    mBalancingOrderContext.CreateBracketOrder( _
                                IIf(mPositionManager.PositionSize < 0, OrderActionBuy, OrderActionSell), _
                                Abs(mPositionManager.PositionSize), _
                                BracketEntryTypes.BracketEntryTypeLimit, _
                                lEntryPrice, _
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
    
    gLogBracketOrderMessage "Close positions at limit: lEntryPrice", _
                    mPositionManager.DataSource, _
                    lContract, _
                    mBalancingBracketOrder.Key, _
                    mPositionManager.IsSimulated, _
                    Me
Else
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
    
    gLogBracketOrderMessage "Close positions at market", _
                    mPositionManager.DataSource, _
                    lContract, _
                    mBalancingBracketOrder.Key, _
                    mPositionManager.IsSimulated, _
                    Me
End If

mBalancingOrderContext.ExecuteBracketOrderEx mBalancingBracketOrder, False, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




