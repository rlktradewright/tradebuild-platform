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
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements CollectionChangeListener
Implements IOrderPersistenceServiceConsumer
Implements IOrderRecoveryServiceConsumer
Implements StateChangeListener

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

Private mOrderPersistenceServiceProvider            As IOrderPersistenceServiceProvider

Private mRecovered                                  As Boolean

Private mOrderPlexes                                As OrderPlexes

Private mAwaitedTickersCollection                   As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mOrderPlexes = New OrderPlexes
Set mAwaitedTickersCollection = New Collection
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As TWUtilities30.CollectionChangeEvent)
Dim tckr As Ticker
If ev.changeType = CollItemAdded Then processStartedTicker ev.affectedItem
End Sub

'@================================================================================
' IOrderPersistenceServiceConsumer Interface Members
'@================================================================================

Private Sub IOrderPersistenceServiceConsumer_PersistedOrderInfo(ByRef orderInfo As String)

Const ProcName As String = "IOrderPersistenceServiceConsumer_PersistedOrderInfo"
Dim failpoint As String
On Error GoTo Err

Dim obj As Object
ParseJSON orderInfo, obj

If TypeOf obj Is OrderPlexSummary Then
    Dim ops As OrderPlexSummary
    Set ops = obj
    
    processOrderPlexSummary ops
Else

End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub IOrderPersistenceServiceConsumer_PersistedOrderInfoEnd()

Const ProcName As String = "IOrderPersistenceServiceConsumer_PersistedOrderInfoEnd"
Dim failpoint As String
On Error GoTo Err

'!!!!!!!!!!!!!!!!!!!!!!! more here !!!!!!!!!!!!!!!!!!!!!!!!!!!!!

mOrderPersistenceServiceProvider.StartSession

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' IOrderRecoveryServiceConsumer Interface Members
'@================================================================================

Private Sub IOrderRecoveryServiceConsumer_OrderReport( _
                ByVal pOrderReport As TradeBuildSP.IOrderReport)

Const ProcName As String = "IOrderRecoveryServiceConsumer_OrderReport"
Dim failpoint As Long
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
        ProjectName & "." & ModuleName & ":" & ProcName, _
        "Not implemented"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As TWUtilities30.StateChangeEvent)
Dim lTicker As Ticker
If ev.State = TickerStates.TickerStateReady Or ev.State = TickerStates.TickerStateRunning Then
    Set lTicker = ev.source
    processStartedTicker lTicker
    lTicker.RemoveStateChangeListener Me
End If
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let OrderPersistenceServiceProvider( _
                ByVal value As IOrderPersistenceServiceProvider)
Set mOrderPersistenceServiceProvider = value
End Property

Friend Property Get OrderPersistenceServiceProvider() As IOrderPersistenceServiceProvider
Set OrderPersistenceServiceProvider = mOrderPersistenceServiceProvider
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub RecoverOrders()
Const ProcName As String = "RecoverOrders"
Dim failpoint As String
On Error GoTo Err

If mRecovered Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "RecoverOrders has already been called and may not be called again"
End If

If mOrderPersistenceServiceProvider Is Nothing Then
    setupDefaultOrderPersistenceSP
End If

mOrderPersistenceServiceProvider.RequestOrderInfo
mRecovered = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addOrderPlexToCollectionWaitingForTicker( _
                ByVal op As orderPlex, _
                ByVal pKey As String, _
                ByVal pTickers As Tickers)
Dim waitingOPs As Collection

Const ProcName As String = "addOrderPlexToCollectionWaitingForTicker"
Dim failpoint As String
On Error GoTo Err

pTickers.AddCollectionChangeListener Me

On Error Resume Next
Set waitingOPs = mAwaitedTickersCollection.Item(pKey)
On Error GoTo Err

If waitingOPs Is Nothing Then
    Set waitingOPs = New Collection
    mAwaitedTickersCollection.Add waitingOPs, pKey
End If

waitingOPs.Add op

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Function CreateOrderPlex(ByVal ops As OrderPlexSummary) As Boolean
Dim ws As Workspace
Dim op As orderPlex

Const ProcName As String = "CreateOrderPlex"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set ws = gTB.WorkSpaces.Item(ops.WorkspaceName)
On Error GoTo Err

If ws Is Nothing Then Exit Function

Set op = New orderPlex
op.Key = ops.Key

setOrderPlexTicker op, ops

mOrderPlexes.Add op

CreateOrderPlex = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function orderPlexExists(ByVal ops As OrderPlexSummary) As Boolean
Const ProcName As String = "orderPlexExists"
Dim failpoint As String
On Error GoTo Err

orderPlexExists = mOrderPlexes.Contains(ops.Key)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub processOrderPlexSummary(ByVal ops As OrderPlexSummary)
Const ProcName As String = "processOrderPlexSummary"
Dim failpoint As String
On Error GoTo Err

If Not orderPlexExists(ops) Then
    If CreateOrderPlex(ops) Then updateOrderPlex ops
Else
    updateOrderPlex ops
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processStartedTicker(ByVal pTicker As Ticker)
Dim waitingOPs As Collection
Dim op As orderPlex

Const ProcName As String = "processStartedTicker"
Dim failpoint As String

On Error GoTo Err

On Error Resume Next
Set waitingOPs = mAwaitedTickersCollection.Item(pTicker.Key)
On Error GoTo Err

If Not waitingOPs Is Nothing Then
    
    If pTicker.State <> TickerStateReady And pTicker.State <> TickerStateRunning Then
        pTicker.AddStateChangeListener Me
    Else
    
        For Each op In waitingOPs
            op.Ticker = pTicker
        Next
    
        mAwaitedTickersCollection.Remove pTicker.Key
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub setOrderPlexTicker( _
                ByVal op As orderPlex, _
                ByVal ops As OrderPlexSummary)
Dim ws As Workspace
Dim tckr As Ticker
Const ProcName As String = "setOrderPlexTicker"
Dim failpoint As String
On Error GoTo Err

Set ws = gTB.WorkSpaces.Item(ops.WorkspaceName)

On Error Resume Next
Set tckr = ws.Tickers.Item(ops.TickerKey)
On Error GoTo Err

If Not tckr Is Nothing Then
    op.Ticker = tckr
Else
    ' the ticker has not yet been started but may yet be
    addOrderPlexToCollectionWaitingForTicker op, ops.TickerKey, ws.Tickers
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub setupDefaultOrderPersistenceSP()
Const ProcName As String = "setupDefaultOrderPersistenceSP"
Dim failpoint As String
On Error GoTo Err

gTB.ServiceProviders.Add ProjectName & ".OrderPersistenceSP", True, "RecoveryFilePath=" & gTB.AppSettingsFolder & "\OrderInfo.txt"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub updateOrderPlex(ByVal ops As OrderPlexSummary)
Dim op As orderPlex
Dim ord As Order

Const ProcName As String = "updateOrderPlex"
Dim failpoint As String
On Error GoTo Err

Set op = mOrderPlexes.Item(ops.Key)

op.cancelPrice = ops.cancelPrice
op.CancelTime = ops.CancelTime
If ops.CloseoutOrderId <> "" And op.CloseoutOrder Is Nothing Then
    Set ord = New Order
    ord.Id = ops.CloseoutOrderId
    op.CloseoutOrder = ord
End If
op.CreationTime = ops.CreationTime
op.Description = ops.Description
op.DrawdownNotificationThreshold = ops.DrawdownNotificationThreshold
If ops.EntryOrderId <> "" And op.EntryOrder Is Nothing Then
    Set ord = New Order
    ord.Id = ops.EntryOrderId
    op.EntryOrder = ord
End If
op.EntryReason = ops.EntryReason
op.LossNotificationThreshold = ops.LossNotificationThreshold
op.PreventUnprotectedPosition = ops.PreventUnprotectedPosition
op.ProfitNotificationThreshold = ops.ProfitNotificationThreshold
op.size = ops.size
op.State = ops.State
If ops.StopOrderId <> "" And op.stoporder Is Nothing Then
    Set ord = New Order
    ord.Id = ops.StopOrderId
    op.stoporder = ord
End If
op.StopReason = ops.StopReason
If ops.TargetOrderId <> "" And op.TargetOrder Is Nothing Then
    Set ord = New Order
    ord.Id = ops.TargetOrderId
    op.TargetOrder = ord
End If
op.TargetReason = ops.TargetOrderId

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub


