VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPlexRecoverer"
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

Private Const ModuleName                            As String = "OrderPlexRecoverer"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext
Private mCancel                                     As Boolean

Private mOrderPlex                                  As orderPlex

Private mWorkspace                                  As WorkSpace
Private mTickerKey                                  As String

Private mPositionManager                            As PositionManager
Private mOSSC                                       As IOrderSubmissionSrvcConsumer

Private mOrderContext                               As OrderContext

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As TWUtilities30.CollectionChangeEvent)
Const ProcName As String = "CollectionChangeListener_Change"

On Error GoTo Err

If ev.changeType = CollItemAdded Then processStartedTicker ev.affectedItem

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As TWUtilities30.StateChangeEvent)
Dim lTicker As Ticker
Const ProcName As String = "StateChangeListener_Change"

On Error GoTo Err

If ev.State = TickerStates.TickerStateReady Or ev.State = TickerStates.TickerStateRunning Then
    Set lTicker = ev.source
    lTicker.RemoveStateChangeListener Me
    
    ActivateOrderPlex lTicker
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
mCancel = True
End Sub

Private Sub Task_run()
If mCancel Then
    mTaskContext.Finish Null, True
    Exit Sub
End If

mTaskContext.suspend -1
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Key() As String
Key = mOrderPlex.Key
End Property

Friend Property Get orderPlex() As orderPlex
Set orderPlex = mOrderPlex
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub FinishRecovery()

Const ProcName As String = "FinishRecovery"
On Error GoTo Err

setUnrecoveredOrderAsCancelled mOrderPlex.EntryOrder
setUnrecoveredOrderAsCancelled mOrderPlex.Stoporder
setUnrecoveredOrderAsCancelled mOrderPlex.TargetOrder
setUnrecoveredOrderAsCancelled mOrderPlex.CloseoutOrder

mOrderPlex.CheckComplete
Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Initialise( _
                ByVal pKey As String, _
                ByVal pWorkspaceName As String, _
                ByVal pTickerKey As String)

Const ProcName As String = "Initialise"

On Error GoTo Err

logMessage "Starting recovery of orderPlex", pMsgQualifier:=pKey, pProcName:=ProcName, pLogLevel:=LogLevelDetail

Set mOrderPlex = New orderPlex
mOrderPlex.Key = pKey
mOrderPlex.State = OrderPlexStateSubmitted
Set mWorkspace = gTB.WorkSpaces.Item(pWorkspaceName)
mTickerKey = pTickerKey

setOrderPlexTicker

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Friend Sub NotifyExecutionReport( _
                ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "NotifyExecutionReport"
Dim lOrder As Order

On Error GoTo Err

mOSSC.ExecutionReport pExecutionReport

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub NotifyOrderReport( _
                ByVal pOrderReport As IOrderReport)

Const ProcName As String = "NotifyOrderReport"

On Error GoTo Err

If orderIdIsForEntryOrder(pOrderReport.TradeBuildId) Then
    mOrderPlex.EntryOrder.SyncToOrderReport pOrderReport, mPositionManager.Key
    mOrderPlex.EntryOrder.NeedsRecovery = False
    'mOrderPlex.ResyncEntryOrder
ElseIf orderIdIsForStopOrder(pOrderReport.TradeBuildId) Then
    mOrderPlex.Stoporder.SyncToOrderReport pOrderReport, mPositionManager.Key
    mOrderPlex.Stoporder.NeedsRecovery = False
    'mOrderPlex.ResyncStopOrder
    
    If mOrderPlex.Stoporder.Status = OrderStatusSubmitted Then
        mOrderPlex.EntryOrder.UpdateStatus OrderStatusFilled, mPositionManager.Key
    End If
ElseIf orderIdIsForTargetOrder(pOrderReport.TradeBuildId) Then
    mOrderPlex.TargetOrder.SyncToOrderReport pOrderReport, mPositionManager.Key
    mOrderPlex.TargetOrder.NeedsRecovery = False
    'mOrderPlex.ResyncTargetOrder

    If mOrderPlex.TargetOrder.Status = OrderStatusSubmitted Then
        mOrderPlex.EntryOrder.UpdateStatus OrderStatusFilled, mPositionManager.Key
    End If
ElseIf orderIdIsForCloseoutOrder(pOrderReport.TradeBuildId) Then
    mOrderPlex.CloseoutOrder.SyncToOrderReport pOrderReport, mPositionManager.Key
    mOrderPlex.CloseoutOrder.NeedsRecovery = False
    'mOrderPlex.ResyncCloseoutOrder

    If mOrderPlex.CloseoutOrder.Status = OrderStatusSubmitted Then
        mOrderPlex.EntryOrder.UpdateStatus OrderStatusFilled, mPositionManager.Key
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub UpdateExecution( _
                ByVal pExec As Execution)
Const ProcName As String = "UpdateExecution"

On Error GoTo Err

mOSSC.ExecutionReport pExec

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub UpdateOrder(ByVal os As OrderSummary)
Dim lOrder As Order

Const ProcName As String = "updateOrder"

On Error GoTo Err

If orderIdIsForEntryOrder(os.Id) Then
    Set lOrder = mOrderPlex.EntryOrder
ElseIf orderIdIsForStopOrder(os.Id) Then
    Set lOrder = mOrderPlex.Stoporder
ElseIf orderIdIsForTargetOrder(os.Id) Then
    Set lOrder = mOrderPlex.TargetOrder
ElseIf orderIdIsForCloseoutOrder(os.Id) Then
    Set lOrder = mOrderPlex.CloseoutOrder
End If

lOrder.action = os.action
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Friend Sub UpdateOrderPlex(ByVal ops As OrderPlexSummary)
Const ProcName As String = "UpdateOrderPlex"

On Error GoTo Err

mOrderPlex.SyncToSummary ops

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub ActivateOrderPlex( _
                ByVal pTicker As Ticker)
Const ProcName As String = "ActivateOrderPlex"

On Error GoTo Err

If Not mOrderPlex.EntryOrder Is Nothing Then mOrderPlex.EntryOrder.Ticker = pTicker
If Not mOrderPlex.Stoporder Is Nothing Then mOrderPlex.Stoporder.Ticker = pTicker
If Not mOrderPlex.TargetOrder Is Nothing Then mOrderPlex.TargetOrder.Ticker = pTicker
If Not mOrderPlex.CloseoutOrder Is Nothing Then mOrderPlex.CloseoutOrder.Ticker = pTicker

Set mOrderContext = pTicker.DefaultOrderContext
mOrderContext.ActivateOrderPlex mOrderPlex

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
GLogLogger.Log pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Function orderIdIsForCloseoutOrder( _
                ByVal pId As String) As Boolean
Const ProcName As String = "orderIdIsForCloseoutOrder"

On Error GoTo Err

If mOrderPlex.CloseoutOrder Is Nothing Then Exit Function
orderIdIsForCloseoutOrder = (pId = mOrderPlex.CloseoutOrder.Id)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function orderIdIsForEntryOrder( _
                ByVal pId As String) As Boolean
Const ProcName As String = "orderIdIsForEntryOrder"

On Error GoTo Err

If mOrderPlex.EntryOrder Is Nothing Then Exit Function
orderIdIsForEntryOrder = (pId = mOrderPlex.EntryOrder.Id)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function orderIdIsForStopOrder( _
                ByVal pId As String) As Boolean
Const ProcName As String = "orderIdIsForStopOrder"

On Error GoTo Err

If mOrderPlex.Stoporder Is Nothing Then Exit Function
orderIdIsForStopOrder = (pId = mOrderPlex.Stoporder.Id)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function orderIdIsForTargetOrder( _
                ByVal pId As String) As Boolean
Const ProcName As String = "orderIdIsForTargetOrder"

On Error GoTo Err

If mOrderPlex.TargetOrder Is Nothing Then Exit Function
orderIdIsForTargetOrder = (pId = mOrderPlex.TargetOrder.Id)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function orderPlexActive() As Boolean
Const ProcName As String = "orderPlexActive"

On Error GoTo Err

orderPlexActive = (Not mOrderPlex.Ticker Is Nothing)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub processStartedTicker(ByVal pTicker As Ticker)
Const ProcName As String = "processStartedTicker"


On Error GoTo Err

If pTicker.Key = mTickerKey Then
    If pTicker.State <> TickerStateReady And pTicker.State <> TickerStateRunning Then
        pTicker.AddStateChangeListener Me
    Else
        ActivateOrderPlex pTicker
        mWorkspace.Tickers.RemoveCollectionChangeListener Me
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setOrderPlexTicker()
Dim tckr As Ticker
Const ProcName As String = "setOrderPlexTicker"

On Error GoTo Err

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
    ActivateOrderPlex tckr
Else
    ' the ticker has been started but is not yet in the right state
    Set mPositionManager = tckr.PositionManager
    tckr.AddStateChangeListener Me
End If

mPositionManager.StartRecovery

Set mOSSC = mPositionManager
mOrderPlex.PositionManager = mPositionManager
mOrderPlex.OrderContext = mPositionManager.OrderContexts.DefaultOrderContext
Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setUnrecoveredOrderAsCancelled(ByVal pOrder As Order)
If pOrder Is Nothing Then Exit Sub
If pOrder.NeedsRecovery Then
    pOrder.Status = OrderStatusCancelled
    pOrder.NeedsRecovery = False
End If
End Sub

