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
Private mBalancingBracketOrderPlaced                As Boolean

Private mOrderContextsWasEnabled                    As Boolean

Private mPositionManager                            As PositionManager
Private mContract                                   As IContract

Private mLastExecutionReport                        As IExecutionReport

Private mCloseoutMode                               As CloseoutModes
Private mPriceSpec                                  As PriceSpecifier

Private mBracketOrderFutureBuilder                  As FutureBuilder

Private mNumberOfActiveOrderContexts                As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mBracketOrderFutureBuilder = New FutureBuilder
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

If ev.changeType <> OrderContextChangeTypes.OrderContextActiveStateChanged Then Exit Sub

Dim lOrderContext As OrderContext
Set lOrderContext = ev.Source

If lOrderContext.IsActive Then Exit Sub

mNumberOfActiveOrderContexts = mNumberOfActiveOrderContexts - 1
lOrderContext.RemoveChangeListener Me

If mNumberOfActiveOrderContexts = 0 Then Assert Not mOrderContexts.HasActiveBracketOrders, _
                                                "Order context still has active bracket orders"

' If we get here, all bracket orders must be closed.

ensureAllBracketOrdersClosedOut

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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
    mBracketOrderFutureBuilder.Value = Nothing
    mBracketOrderFutureBuilder.Complete
    RaiseEvent PositionClosed
    Exit Sub
End If

Set mOrderContexts = mPositionManager.OrderContexts
Set mContract = mPositionManager.ContractFuture.Value

mOrderContextsWasEnabled = mOrderContexts.Enabled
mOrderContexts.Enabled = False  ' prevent any further orders being placed
                                ' while the position is being closed

If Not cancelOrders Then ensureAllBracketOrdersClosedOut

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mBalancingBracketOrder Event Handlers
'@================================================================================

Private Sub mBalancingBracketOrder_Completed()
Const ProcName As String = "mBalancingBracketOrder_Completed"
On Error GoTo Err

mBalancingBracketOrder.SetSize DecimalZero

' adjust the other BracketOrders to take account of the settling Order

If mLastExecutionReport Is Nothing Then
    ' the balancing bracket order has failed due to some error
    doFinalActions
    RaiseEvent PositionNotClosed(mBalancingBracketOrder.EntryOrder.ErrorMessage)
    Exit Sub
End If

Dim lActualCloseoutOrder As Order
Set lActualCloseoutOrder = mBalancingBracketOrder.CloseoutOrder

Dim oc As OrderContext
Dim i As Long
For Each oc In mOrderContexts
    If Not oc Is mBalancingOrderContext Then
        Dim lBracketOrder As IBracketOrder
        For Each lBracketOrder In oc.ActiveBracketOrders
            doSettlingCloseout oc, lBracketOrder
        Next
        For Each lBracketOrder In oc.ClosedBracketOrders
            doSettlingCloseout oc, lBracketOrder
        Next
    End If
Next

doFinalActions
RaiseEvent PositionClosed

Exit Sub

Err:
GOrders.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBalancingBracketOrder_Execution(ev As ExecutionEventData)
Const ProcName As String = "mBalancingBracketOrder_Execution"
On Error GoTo Err

Set mLastExecutionReport = ev.ExecutionReport

Exit Sub

Err:
GOrders.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mBalancingOrderContext Event Handlers
'@================================================================================

Private Sub mBalancingOrderContext_Change(ev As ChangeEventData)
Const ProcName As String = "mBalancingOrderContext_Change"
On Error GoTo Err

If ev.changeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub
If mBalancingBracketOrderPlaced Then Exit Sub
If mPositionManager.PositionSize = 0 Then Exit Sub


If mBalancingBracketOrder Is Nothing Then
    placeBalancingOrder
ElseIf mBalancingOrderContext.IsReadyForBracketOrder(mBalancingBracketOrder) Then
    mBalancingOrderContext.ExecuteBracketOrderEx mBalancingBracketOrder, False, False
    mBalancingBracketOrderPlaced = True
Else
    GOrderUtils.Log mBalancingBracketOrder.Key & _
            ": Balancing order context: " & _
            mBalancingOrderContext.NotReadyMessage, _
            ProcName, ModuleName
End If

Exit Sub

Err:
GOrders.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function ClosePositions( _
                ByVal pPositionManager As PositionManager, _
                ByVal pCloseoutMode As CloseoutModes, _
                ByVal pPriceSpec As PriceSpecifier) As IFuture
Const ProcName As String = "ClosePositions"
On Error GoTo Err

Set mPositionManager = pPositionManager
mCloseoutMode = pCloseoutMode
Set mPriceSpec = pPriceSpec

Set ClosePositions = mBracketOrderFutureBuilder.Future

DeferAction Me

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

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
GOrders.HandleUnexpectedError ProcName, ModuleName
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
        mNumberOfActiveOrderContexts = mNumberOfActiveOrderContexts + 1
    End If
Next

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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
                    CreateBoxedDecimal(Abs(mPositionManager.PositionSize)), _
                    mBalancingOrderContext.CreateEntryOrder( _
                        OrderTypeLimit, _
                        pPriceSpec _
                    ))

LogBracketOrderMessage "Close position at limit: " & gPriceSpecifierToString(mPriceSpec, pContract), _
                mPositionManager.DataSource, _
                pContract, _
                lBo.Key, _
                mPositionManager.IsSimulated, _
                Me

Set createLimitBalancingOrder = lBo

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function createMarketBalancingOrder( _
                ByVal pContract As IContract) As IBracketOrder
Const ProcName As String = "createMarketBalancingOrder"
On Error GoTo Err

Dim lBo As BracketOrder

Set lBo = mBalancingOrderContext.CreateBracketOrder( _
                        IIf(mPositionManager.PositionSize < 0, OrderActionBuy, OrderActionSell), _
                        CreateBoxedDecimal(Abs(mPositionManager.PositionSize)), _
                        mBalancingOrderContext.CreateEntryOrder( _
                            OrderTypeMarket) _
                        )

LogBracketOrderMessage "Close position at market", _
                mPositionManager.DataSource, _
                pContract, _
                lBo.Key, _
                mPositionManager.IsSimulated, _
                Me

Set createMarketBalancingOrder = lBo

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub doFinalActions()
Const ProcName As String = "doFinalActions"
On Error GoTo Err

mOrderContexts.Enabled = mOrderContextsWasEnabled
If mPositionManager.EndingSession Then
    GOrderUtils.LogMoneyManagement "Ending Session", mPositionManager.IsSimulated, Me
End If
Set mBalancingOrderContext = Nothing

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doMarkToMaketCloseout( _
                ByVal pOrderContext As OrderContext, _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pTradePrice As Double)
Const ProcName As String = "doMarkToMaketCloseout"
On Error GoTo Err

Dim lExec As Execution: Set lExec = pOrderContext.NewExecution(True)
lExec.GroupName = pOrderContext.GroupName
lExec.RecoveryContextName = pOrderContext.RecoveryContextName
lExec.Quantity = CreateBoxedDecimal(Abs(pBracketOrder.Size))
If pBracketOrder.Size > 0 Then
    lExec.Action = OrderActionSell
Else
    lExec.Action = OrderActionBuy
End If
lExec.FillingExchange = "*"
lExec.FillTime = pOrderContext.Clock.Timestamp
lExec.Id = "MarkToMarket:" & GenerateGUIDString
lExec.Price = pTradePrice
lExec.SecurityName = GOrderUtils.GetContractName(mContract.Specifier)
lExec.TimezoneName = mContract.TimezoneName

pBracketOrder.NotifyCloseout lExec

lExec.SaveRecoveryInfo

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doSettlingCloseout( _
                ByVal pOrderContext As OrderContext, _
                ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "doSettlingCloseout"
On Error GoTo Err

If pBracketOrder.Size = 0 Then Exit Sub
    
' NB: we don't save recovery data for this execution
' because it will simply be recreated at recovery

Dim lExec As Execution
Set lExec = pOrderContext.NewExecution(True)
lExec.GroupName = pOrderContext.GroupName
lExec.RecoveryContextName = pOrderContext.RecoveryContextName
lExec.SyncToExecutionReport mLastExecutionReport

' give this execution a unique id
Static i As Long
i = i + 1
lExec.Id = lExec.Id & "{" & i & "}"

lExec.Quantity = CreateBoxedDecimal(Abs(pBracketOrder.Size))
If pBracketOrder.Size > 0 Then
    lExec.Action = OrderActionSell
Else
    lExec.Action = OrderActionBuy
End If

pBracketOrder.NotifyCloseout lExec

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ensureAllBracketOrdersClosedOut()
Const ProcName As String = "ensureAllBracketOrdersClosedOut"
On Error GoTo Err

Set mBalancingOrderContext = mOrderContexts.BalancingOrderContext
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub markToMarket()
Const ProcName As String = "markToMarket"
On Error GoTo Err

Dim lNoTickAvailable As Boolean
Dim lTradePrice As Double

' closeout non-zero-size bracket orders marking them to mnarket

If mPositionManager.DataSource.HasCurrentTick(TickTypeTrade) Then
    lTradePrice = mPositionManager.DataSource.CurrentTick(TickTypeTrade).Price
Else
    lNoTickAvailable = True
End If

Dim oc As OrderContext
For Each oc In mOrderContexts
    If Not oc Is mBalancingOrderContext Then
        Dim lBracketOrder As IBracketOrder
        For Each lBracketOrder In oc.ActiveBracketOrders
            If lBracketOrder.Size <> 0 Then
                If lNoTickAvailable Then
                    RaiseEvent PositionNotClosed("No trade tick available")
                    Exit Sub
                End If
                
                doMarkToMaketCloseout oc, lBracketOrder, lTradePrice
            End If
        Next
        For Each lBracketOrder In oc.ClosedBracketOrders
            If lBracketOrder.Size <> 0 Then
                If lNoTickAvailable Then
                    RaiseEvent PositionNotClosed("No trade tick available")
                    Exit Sub
                End If
                
                doMarkToMaketCloseout oc, lBracketOrder, lTradePrice
            End If
        Next
    End If
Next

doFinalActions
RaiseEvent PositionClosed

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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

mBracketOrderFutureBuilder.Value = mBalancingBracketOrder
mBracketOrderFutureBuilder.Complete

If mBalancingOrderContext.IsReadyForBracketOrder(mBalancingBracketOrder) Then
    mBalancingOrderContext.ExecuteBracketOrderEx mBalancingBracketOrder, False, False
    mBalancingBracketOrderPlaced = True
Else
    GOrderUtils.Log mBalancingBracketOrder.Key & _
            ": Balancing order context: " & _
            mBalancingOrderContext.NotReadyMessage, _
            ProcName, ModuleName
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub




