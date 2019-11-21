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

Private mCloseoutMode                               As CloseoutModes
Private mPriceSpec                                  As PriceSpecifier

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
RaiseEvent PositionClosed

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
If mBalancingBracketOrder Is Nothing Then
    placeBalancingOrder
ElseIf mBalancingOrderContext.IsReadyForBracketOrder(mBalancingBracketOrder) Then
    mBalancingOrderContext.ExecuteBracketOrderEx mBalancingBracketOrder, False, False
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
                ByVal pCloseoutMode As CloseoutModes, _
                ByVal pPriceSpec As PriceSpecifier)
Const ProcName As String = "ClosePositions"
On Error GoTo Err

Set mPositionManager = pPositionManager
mCloseoutMode = pCloseoutMode
Set mPriceSpec = pPriceSpec

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

Private Function createLimitBalancingOrder( _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pContract As IContract) As IBracketOrder
Const ProcName As String = "createLimitBalancingOrder"
On Error GoTo Err

Dim lBo As BracketOrder

If Not mBalancingOrderContext.IsReadyForPriceSpecifier(pPriceSpec) Then Exit Function

Set lBo = mBalancingOrderContext.CreateBracketOrder( _
                    IIf(mPositionManager.PositionSize < 0, OrderActionBuy, OrderActionSell), _
                    Abs(mPositionManager.PositionSize), _
                    mBalancingOrderContext.CreateEntryOrder( _
                        OrderTypeLimit, _
                        pPriceSpec _
                    ))

gLogBracketOrderMessage "Close position at limit: " & gPriceSpecifierToString(mPriceSpec, pContract), _
                mPositionManager.DataSource, _
                pContract, _
                lBo.Key, _
                mPositionManager.IsSimulated, _
                Me

Set createLimitBalancingOrder = lBo

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createMarketBalancingOrder( _
                ByVal pContract As IContract) As IBracketOrder
Const ProcName As String = "createMarketBalancingOrder"
On Error GoTo Err

Dim lBo As BracketOrder

Set lBo = mBalancingOrderContext.CreateBracketOrder( _
                        IIf(mPositionManager.PositionSize < 0, OrderActionBuy, OrderActionSell), _
                        Abs(mPositionManager.PositionSize), _
                        mBalancingOrderContext.CreateEntryOrder( _
                            OrderTypeMarket) _
                        )

gLogBracketOrderMessage "Close position at market", _
                mPositionManager.DataSource, _
                pContract, _
                lBo.Key, _
                mPositionManager.IsSimulated, _
                Me

Set createMarketBalancingOrder = lBo

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
Set mBalancingOrderContext = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ensureAllBracketOrdersClosedOut()
Const ProcName As String = "ensureAllBracketOrdersClosedOut"
On Error GoTo Err

If mPositionManager.PositionSize <> 0 Then
    Set mBalancingOrderContext = mOrderContexts.BalancingOrderContext
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

If Not mPositionManager.DataSource.HasCurrentTick(TickTypeTrade) Then
    RaiseEvent PositionNotClosed("No trade tick available")
    Exit Sub
End If

Dim lTradePrice As Double: lTradePrice = mPositionManager.DataSource.CurrentTick(TickTypeTrade).Price

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
                lExec.FillingExchange = "*"
                lExec.FillTime = mClock.Timestamp
                lExec.Id = "MarkToMarket:" & GenerateGUIDString
                lExec.Price = lTradePrice
                lExec.SecurityName = mContract.Specifier.LocalSymbol & "@" & mContract.Specifier.Exchange
                lExec.TimezoneName = mContract.TimezoneName
                
                lBracketOrder.NotifyCloseout lExec
            
                lExec.SaveRecoveryInfo
            End If
        Next
    End If
Next

doFinalActions
RaiseEvent PositionClosed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub placeBalancingOrder()
Const ProcName As String = "placeBalancingOrder"
On Error GoTo Err

Dim lContract As IContract
Set lContract = mPositionManager.DataSource.ContractFuture.Value

If mCloseoutMode = CloseoutModeMarket Then
    Set mBalancingBracketOrder = createMarketBalancingOrder(lContract)
ElseIf mCloseoutMode = CloseoutModeLimit Then
    Set mBalancingBracketOrder = createLimitBalancingOrder(mPriceSpec, lContract)
Else
End If

If mBalancingOrderContext.IsReadyForBracketOrder(mBalancingBracketOrder) Then
    mBalancingOrderContext.ExecuteBracketOrderEx mBalancingBracketOrder, False, False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




