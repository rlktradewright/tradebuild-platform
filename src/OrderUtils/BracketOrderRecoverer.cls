VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderRecoverer"
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

Implements CollectionChangeListener
Implements StateChangeListener
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

Private Const ModuleName                            As String = "BracketOrderRecoverer"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext
Private mCancel                                     As Boolean

Private mBracketOrder                               As BracketOrder

Private mTickerKey                                  As String

Private mPositionManager                            As PositionManager
Private mOSL                                        As IOrderSubmissionListener
    
Private mOrderContext                               As OrderContext

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As CollectionChangeEventData)
Const ProcName As String = "CollectionChangeListener_Change"

On Error GoTo Err

If ev.changeType = CollItemAdded Then processStartedTicker ev.AffectedItem

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Dim lTicker As Ticker
Const ProcName As String = "StateChangeListener_Change"

On Error GoTo Err

If ev.State = TickerStates.TickerStateReady Or ev.State = TickerStates.TickerStateRunning Then
    Set lTicker = ev.Source
    lTicker.RemoveStateChangeListener Me
    
    ActivateBracketOrder lTicker
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
mCancel = True
End Sub

Private Sub Task_Run()
If mCancel Then
    mTaskContext.Finish Null, True
    Exit Sub
End If

mTaskContext.Suspend -1
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Key() As String
Key = mBracketOrder.Key
End Property

Friend Property Get BracketOrder() As BracketOrder
Set BracketOrder = mBracketOrder
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub FinishRecovery()

Const ProcName As String = "FinishRecovery"
On Error GoTo Err

setUnrecoveredOrderAsCancelled mBracketOrder.EntryOrder
setUnrecoveredOrderAsCancelled mBracketOrder.Stoporder
setUnrecoveredOrderAsCancelled mBracketOrder.TargetOrder
setUnrecoveredOrderAsCancelled mBracketOrder.CloseoutOrder

mBracketOrder.CheckComplete
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pKey As String, _
                ByVal pWorkspaceName As String, _
                ByVal pTickerKey As String)

Const ProcName As String = "Initialise"

On Error GoTo Err

gLog "Starting recovery of BracketOrder", pMsgQualifier:=pKey, pModName:=ModuleName, pProcName:=ProcName, pLogLevel:=LogLevelDetail

Set mBracketOrder = New BracketOrder
mBracketOrder.Key = pKey
mBracketOrder.State = BracketOrderStateSubmitted
Set mWorkspace = gTB.WorkSpaces.Item(pWorkspaceName)
mTickerKey = pTickerKey

setBracketOrderTicker

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Friend Sub NotifyExecutionReport( _
                ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "NotifyExecutionReport"
Dim lOrder As Order

On Error GoTo Err

mOSL.NotifyExecutionReport pExecutionReport

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyOrderReport( _
                ByVal pOrderReport As IOrderReport)

Const ProcName As String = "NotifyOrderReport"

On Error GoTo Err

If orderIdIsForEntryOrder(pOrderReport.TradeBuildId) Then
    mBracketOrder.EntryOrder.SyncToOrderReport pOrderReport, mPositionManager.Key
    mBracketOrder.EntryOrder.NeedsRecovery = False
    'mBracketOrder.ResyncEntryOrder
ElseIf orderIdIsForStopOrder(pOrderReport.TradeBuildId) Then
    mBracketOrder.Stoporder.SyncToOrderReport pOrderReport, mPositionManager.Key
    mBracketOrder.Stoporder.NeedsRecovery = False
    'mBracketOrder.ResyncStopOrder
    
    If mBracketOrder.Stoporder.Status = OrderStatusSubmitted Then
        mBracketOrder.EntryOrder.UpdateStatus OrderStatusFilled, mPositionManager.Key
    End If
ElseIf orderIdIsForTargetOrder(pOrderReport.TradeBuildId) Then
    mBracketOrder.TargetOrder.SyncToOrderReport pOrderReport, mPositionManager.Key
    mBracketOrder.TargetOrder.NeedsRecovery = False
    'mBracketOrder.ResyncTargetOrder

    If mBracketOrder.TargetOrder.Status = OrderStatusSubmitted Then
        mBracketOrder.EntryOrder.UpdateStatus OrderStatusFilled, mPositionManager.Key
    End If
ElseIf orderIdIsForCloseoutOrder(pOrderReport.TradeBuildId) Then
    mBracketOrder.CloseoutOrder.SyncToOrderReport pOrderReport, mPositionManager.Key
    mBracketOrder.CloseoutOrder.NeedsRecovery = False
    'mBracketOrder.ResyncCloseoutOrder

    If mBracketOrder.CloseoutOrder.Status = OrderStatusSubmitted Then
        mBracketOrder.EntryOrder.UpdateStatus OrderStatusFilled, mPositionManager.Key
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub UpdateExecution( _
                ByVal pExec As IExecutionReport)
Const ProcName As String = "UpdateExecution"

On Error GoTo Err

mOSL.NotifyExecutionReport pExec

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub UpdateOrder(ByVal os As OrderSummary)
Dim lOrder As Order

Const ProcName As String = "updateOrder"

On Error GoTo Err

If orderIdIsForEntryOrder(os.Id) Then
    Set lOrder = mBracketOrder.EntryOrder
ElseIf orderIdIsForStopOrder(os.Id) Then
    Set lOrder = mBracketOrder.Stoporder
ElseIf orderIdIsForTargetOrder(os.Id) Then
    Set lOrder = mBracketOrder.TargetOrder
ElseIf orderIdIsForCloseoutOrder(os.Id) Then
    Set lOrder = mBracketOrder.CloseoutOrder
End If

lOrder.Action = os.Action
lOrder.AllOrNone = os.AllOrNone
lOrder.AveragePrice = os.AveragePrice
lOrder.BlockOrder = os.BlockOrder
lOrder.BrokerId = os.BrokerId
lOrder.DiscretionaryAmount = os.DiscretionaryAmount
lOrder.DisplaySize = os.DisplaySize
lOrder.ETradeOnly = os.ETradeOnly
lOrder.FillTime = os.FillTime
lOrder.FirmQuoteOnly = os.FirmQuoteOnly
lOrder.GoodAfterTime = os.GoodAfterTime
lOrder.GoodAfterTimeTZ = os.GoodAfterTimeTZ
lOrder.GoodTillDate = os.GoodTillDate
lOrder.GoodTillDateTZ = os.GoodTillDateTZ
lOrder.Hidden = os.Hidden
lOrder.Id = os.Id
lOrder.IgnoreRegularTradingHours = os.IgnoreRegularTradingHours
lOrder.LastFillPrice = os.LastFillPrice
lOrder.LimitPrice = os.LimitPrice
lOrder.MinimumQuantity = os.MinimumQuantity
lOrder.NbboPriceCap = os.NbboPriceCap
lOrder.Offset = os.Offset
lOrder.OrderType = os.OrderType
lOrder.Origin = os.Origin
lOrder.OriginatorRef = os.OriginatorRef
lOrder.OverrideConstraints = os.OverrideConstraints
lOrder.PercentOffset = os.PercentOffset
lOrder.ProviderProperties = os.ProviderProperties
lOrder.Quantity = os.Quantity
lOrder.QuantityFilled = os.QuantityFilled
lOrder.QuantityRemaining = os.QuantityRemaining
lOrder.SettlingFirm = os.SettlingFirm
lOrder.Status = os.Status
lOrder.StopTriggerMethod = os.StopTriggerMethod
lOrder.SweepToFill = os.SweepToFill
lOrder.TimeInForce = os.TimeInForce
lOrder.TriggerPrice = os.TriggerPrice

lOrder.NeedsRecovery = lOrder.IsActive
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Friend Sub UpdateBracketOrder(ByVal ops As BracketOrderSummary)
Const ProcName As String = "UpdateBracketOrder"

On Error GoTo Err

mBracketOrder.SyncToSummary ops

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub ActivateBracketOrder( _
                ByVal pTicker As Ticker)
Const ProcName As String = "ActivateBracketOrder"

On Error GoTo Err

If Not mBracketOrder.EntryOrder Is Nothing Then mBracketOrder.EntryOrder.Ticker = pTicker
If Not mBracketOrder.Stoporder Is Nothing Then mBracketOrder.Stoporder.Ticker = pTicker
If Not mBracketOrder.TargetOrder Is Nothing Then mBracketOrder.TargetOrder.Ticker = pTicker
If Not mBracketOrder.CloseoutOrder Is Nothing Then mBracketOrder.CloseoutOrder.Ticker = pTicker

Set mOrderContext = pTicker.DefaultOrderContext
mOrderContext.ActivateBracketOrder mBracketOrder

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function createOrder(ByVal Id As String)
Dim lOrder As Order
Const ProcName As String = "createOrder"

On Error GoTo Err

Set lOrder = New Order
lOrder.Id = Id
Set createOrder = lOrder

mPositionManager.Orders.Add lOrder

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function orderIdIsForCloseoutOrder( _
                ByVal pId As String) As Boolean
Const ProcName As String = "orderIdIsForCloseoutOrder"

On Error GoTo Err

If mBracketOrder.CloseoutOrder Is Nothing Then Exit Function
orderIdIsForCloseoutOrder = (pId = mBracketOrder.CloseoutOrder.Id)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function orderIdIsForEntryOrder( _
                ByVal pId As String) As Boolean
Const ProcName As String = "orderIdIsForEntryOrder"

On Error GoTo Err

If mBracketOrder.EntryOrder Is Nothing Then Exit Function
orderIdIsForEntryOrder = (pId = mBracketOrder.EntryOrder.Id)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function orderIdIsForStopOrder( _
                ByVal pId As String) As Boolean
Const ProcName As String = "orderIdIsForStopOrder"

On Error GoTo Err

If mBracketOrder.Stoporder Is Nothing Then Exit Function
orderIdIsForStopOrder = (pId = mBracketOrder.Stoporder.Id)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function orderIdIsForTargetOrder( _
                ByVal pId As String) As Boolean
Const ProcName As String = "orderIdIsForTargetOrder"

On Error GoTo Err

If mBracketOrder.TargetOrder Is Nothing Then Exit Function
orderIdIsForTargetOrder = (pId = mBracketOrder.TargetOrder.Id)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function BracketOrderActive() As Boolean
Const ProcName As String = "BracketOrderActive"

On Error GoTo Err

BracketOrderActive = (Not mBracketOrder.Ticker Is Nothing)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub processStartedTicker(ByVal pTicker As Ticker)
Const ProcName As String = "processStartedTicker"
On Error GoTo Err

If pTicker.Key = mTickerKey Then
    If pTicker.State <> TickerStateReady And pTicker.State <> TickerStateRunning Then
        pTicker.AddStateChangeListener Me
    Else
        ActivateBracketOrder pTicker
        mWorkspace.Tickers.RemoveCollectionChangeListener Me
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setBracketOrderTicker()
Const ProcName As String = "setBracketOrderTicker"
On Error GoTo Err

Dim tckr As Ticker

On Error Resume Next
Set tckr = mWorkspace.Tickers.Item(mTickerKey)
On Error GoTo Err

If tckr Is Nothing Then
    ' the ticker has not yet been started but may yet be
    mWorkspace.Tickers.AddCollectionChangeListener Me
    
    Set mPositionManager = gGetPositionManager(mTickerKey)
    If mPositionManager Is Nothing Then
        Set mPositionManager = gCreatePositionManager(mTickerKey, mWorkspace)
    End If
ElseIf tckr.State = TickerStateReady Or tckr.State = TickerStateRunning Then
    Set mPositionManager = tckr.PositionManager
    ActivateBracketOrder tckr
Else
    ' the ticker has been started but is not yet in the right state
    Set mPositionManager = tckr.PositionManager
    tckr.AddStateChangeListener Me
End If

mPositionManager.StartRecovery

Set mOSL = mPositionManager
mBracketOrder.PositionManager = mPositionManager
mBracketOrder.OrderContext = mPositionManager.OrderContexts.DefaultOrderContext
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setUnrecoveredOrderAsCancelled(ByVal pOrder As Order)
If pOrder Is Nothing Then Exit Sub
If pOrder.NeedsRecovery Then
    pOrder.Status = OrderStatusCancelled
    pOrder.NeedsRecovery = False
End If
End Sub

