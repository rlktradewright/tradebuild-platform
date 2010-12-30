VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPlexRecoveryCtlr"
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

Implements IOrderPersistenceServiceConsumer
Implements IOrderRecoveryServiceConsumer
Implements TaskCompletionListener

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

Private Const ModuleName                            As String = "OrderPlexRecoveryCtlr"

Private Const OrderInfoDelete                       As String = "DELETE"
Private Const OrderInfoData                         As String = "DATA"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderPersistenceServiceProvider            As IOrderPersistenceServiceProvider
Private mOrderSubmissionSrvcProvider                As IOrderSubmissionSrvcProvider

Private mRecovered                                  As Boolean

Private mOrderPlexRecoverers                        As Collection

Private mOrderIdMap                                 As OrderIdMap

Private mOrderIdToOrderPlexRecoverer                As Collection

Private mStartRecoveryFrom                          As Date

Private mOrderPersistenceDataGroups                 As EnumerableCollection

Private mExecutionReportsComplete                   As Boolean
Private mOrderReportsComplete                       As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mOrderPlexRecoverers = New Collection
Set mOrderIdMap = New OrderIdMap
Set mOrderIdToOrderPlexRecoverer = New Collection
Set mOrderPersistenceDataGroups = New EnumerableCollection
End Sub

'@================================================================================
' IOrderPersistenceServiceConsumer Interface Members
'@================================================================================

Private Sub IOrderPersistenceServiceConsumer_PersistedOrderInfo( _
                orderInfo As String)

Const ProcName As String = "IOrderPersistenceServiceConsumer_PersistedOrderInfo"

On Error GoTo Err

Dim opItem As OrderPersistenceItem
ParseJSON orderInfo, opItem

If opItem.ItemType = OrderInfoData Then
    Dim opg As OrderPersistenceDataGroup
    On Error Resume Next
    Set opg = mOrderPersistenceDataGroups.Item(opItem.Key)
    On Error GoTo Err
    
    If opg Is Nothing Then
        Set opg = New OrderPersistenceDataGroup
        mOrderPersistenceDataGroups.Add opg, opItem.Key
    End If
    opg.Add opItem
ElseIf opItem.ItemType = OrderInfoDelete Then
    On Error Resume Next
    mOrderPersistenceDataGroups.Remove opItem.Key
    On Error GoTo Err
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderPersistenceServiceConsumer_PersistedOrderInfoEnd()
Dim opr As OrderPlexRecoverer
Dim opItem As OrderPersistenceItem

Const ProcName As String = "IOrderPersistenceServiceConsumer_PersistedOrderInfoEnd"

On Error GoTo Err

mOrderPersistenceServiceProvider.StartSession gTB.SessionId

Do While getNextOrderPersistenceItem(opItem)
    Dim opItemJSON As String
    opItemJSON = opItem.ToJSON
    logMessage "Processing persisted order info", pMsgQualifier:=opItemJSON, pProcName:=ProcName, pLogLevel:=LogLevelDetail

    If processOrderPersistenceItem(opItem) Then
        mOrderPersistenceServiceProvider.SaveOrderInfo opItemJSON
    End If
Loop

If Not mOrderSubmissionSrvcProvider Is Nothing Then
    mOrderSubmissionSrvcProvider.StartOrderRecovery Me, mStartRecoveryFrom
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' IOrderRecoveryServiceConsumer Interface Members
'@================================================================================

Private Sub IOrderRecoveryServiceConsumer_ExecutionReport(ByVal pExecutionReport As TradeBuildSP.IExecutionReport)

Const ProcName As String = "IOrderRecoveryServiceConsumer_ExecutionReport"

On Error GoTo Err

Dim opr As OrderPlexRecoverer
Dim tbId As String

tbId = mOrderIdMap.GetTradeBuildId(pExecutionReport.OrderBrokerId)

If tbId = "" Then
    ' this can happen if the order recovery info has been deleted, so just
    ' ignore
    Exit Sub
End If

pExecutionReport.OrderTradeBuildId = tbId
Set opr = mOrderIdToOrderPlexRecoverer.Item(pExecutionReport.OrderTradeBuildId)

opr.NotifyExecutionReport pExecutionReport

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderRecoveryServiceConsumer_ExecutionReportsComplete()
Const ProcName As String = "IOrderRecoveryServiceConsumer_ExecutionReportsComplete"
On Error GoTo Err

mExecutionReportsComplete = True
If mOrderReportsComplete Then FinishRecovery

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderRecoveryServiceConsumer_OrderReport( _
                ByVal pOrderReport As TradeBuildSP.IOrderReport)

Const ProcName As String = "IOrderRecoveryServiceConsumer_OrderReport"

On Error GoTo Err

Dim opr As OrderPlexRecoverer

pOrderReport.TradeBuildId = mOrderIdMap.GetTradeBuildId(pOrderReport.BrokerId)
Set opr = mOrderIdToOrderPlexRecoverer.Item(pOrderReport.TradeBuildId)

opr.NotifyOrderReport pOrderReport

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrInvalidProcedureCall Then
    ' this can happen if the order recovery info has been deleted, so just
    ' ignore
    Exit Sub
End If
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IOrderRecoveryServiceConsumer_OrderReportsComplete()
Const ProcName As String = "IOrderRecoveryServiceConsumer_OrderReportsComplete"
On Error GoTo Err

mOrderReportsComplete = True
If mExecutionReportsComplete Then FinishRecovery

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As TWUtilities30.TaskCompletionEventData)
Dim opr As OrderPlexRecoverer
Dim op As orderPlex

Set opr = mOrderPlexRecoverers.Item(ev.cookie)
Set op = opr.orderPlex

mOrderPlexRecoverers.Remove ev.cookie
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

Friend Property Let OrderSubmissionSrvcProvider( _
                ByVal value As IOrderSubmissionSrvcProvider)
Set mOrderSubmissionSrvcProvider = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub DeleteRecoveryInfo( _
                ByRef pKey As String)
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

If Not mOrderPersistenceServiceProvider Is Nothing Then
    Dim opItem As New OrderPersistenceItem
    opItem.ItemType = OrderInfoDelete
    opItem.Key = pKey
    opItem.Timestamp = GetTimestamp
    mOrderPersistenceServiceProvider.SaveOrderInfo opItem.ToJSON
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub RecoverOrders(ByVal sessionName As String)
Const ProcName As String = "RecoverOrders"

On Error GoTo Err

If mRecovered Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "RecoverOrders has already been called and may not be called again"
End If

If mOrderPersistenceServiceProvider Is Nothing Then
    setupDefaultOrderPersistenceSP
End If

logMessage "Persisted order info processing started", pProcName:=ProcName

mOrderPersistenceServiceProvider.RequestOrderInfo sessionName
mRecovered = True

logMessage "Persisted order info processing completed", pProcName:=ProcName

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Friend Sub SaveRecoveryInfo( _
                ByRef pKey As String, _
                ByRef info As Object)
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

If pKey = "" Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "pKey cannot be blank"
End If

If Not mOrderPersistenceServiceProvider Is Nothing Then
    Dim opItem As New OrderPersistenceItem
    opItem.data = info
    opItem.ItemType = OrderInfoData
    opItem.Key = pKey
    opItem.Timestamp = GetTimestamp
    mOrderPersistenceServiceProvider.SaveOrderInfo opItem.ToJSON
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function CreateOrderPlexRecoverer(ByVal ops As OrderPlexSummary) As OrderPlexRecoverer
Dim ws As WorkSpace
Dim opr As OrderPlexRecoverer

Const ProcName As String = "CreateOrderPlexRecoverer"

On Error GoTo Err

On Error Resume Next
Set ws = gTB.WorkSpaces.Item(ops.WorkspaceName)
On Error GoTo Err

If ws Is Nothing Then Exit Function

Set opr = New OrderPlexRecoverer
opr.Initialise ops.Key, ops.WorkspaceName, ops.TickerKey
StartTask(opr, PriorityNormal, , opr.Key).AddTaskCompletionListener Me

mOrderPlexRecoverers.Add opr, opr.Key

Set CreateOrderPlexRecoverer = opr

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub FinishRecovery()
Static lFinished As Boolean

Dim opr As OrderPlexRecoverer

Const ProcName As String = "finishRecovery"
On Error GoTo Err

If lFinished Then Exit Sub
lFinished = True

For Each opr In mOrderPlexRecoverers
    opr.FinishRecovery
Next

Set mOrderPlexRecoverers = Nothing

Dim en As Enumerator
Dim pm As PositionManager
Set en = gGetPositionManagersEnumerator
Do While en.MoveNext
    Set pm = en.Current
    If Not pm.FinishRecovery Then en.Remove
Loop

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function getNextOrderPersistenceDataGroupEnumerator( _
                ByRef en As Enumerator) As Boolean
Static orderPersistenceDataGroupsEn As Enumerator
Dim opg As OrderPersistenceDataGroup

Const ProcName As String = "getNextOrderPersistenceDataGroupEnumerator"


On Error GoTo Err

If orderPersistenceDataGroupsEn Is Nothing Then Set orderPersistenceDataGroupsEn = mOrderPersistenceDataGroups.Enumerator

If Not orderPersistenceDataGroupsEn.MoveNext Then Exit Function

Set opg = orderPersistenceDataGroupsEn.Current
Set en = opg.Enumerator

getNextOrderPersistenceDataGroupEnumerator = True
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getNextOrderPersistenceItem( _
                ByRef pOpItem As OrderPersistenceItem) As Boolean
Static opgEn As Enumerator

Const ProcName As String = "getNextOrderPersistenceItem"

On Error GoTo Err

If opgEn Is Nothing Then
    If Not getNextOrderPersistenceDataGroupEnumerator(opgEn) Then Exit Function
End If

Do While Not opgEn.MoveNext
    If Not getNextOrderPersistenceDataGroupEnumerator(opgEn) Then Exit Function
Loop

Set pOpItem = opgEn.Current

getNextOrderPersistenceItem = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
GLogLogger.Log pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub processExecution(ByVal ex As Execution)
Const ProcName As String = "processExecution"
Dim opr As OrderPlexRecoverer


On Error GoTo Err

Set opr = mOrderIdToOrderPlexRecoverer.Item(ex.OrderTradeBuildId)

opr.UpdateExecution ex

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function processOrderPersistenceItem( _
                ByVal pOpItem As OrderPersistenceItem) As Boolean
Const ProcName As String = "processOrderPersistenceItem"
On Error GoTo Err

processOrderPersistenceItem = True

If TypeOf pOpItem.data Is OrderPlexSummary Then
    processOrderPlexSummary pOpItem.data
ElseIf TypeOf pOpItem.data Is OrderSummary Then
    If Not processOrderSummary(pOpItem.data) Then processOrderPersistenceItem = False
ElseIf TypeOf pOpItem.data Is Execution Then
    processExecution pOpItem.data
ElseIf TypeOf pOpItem.data Is PositionManagerSummary Then
    processPositionManagerSummary pOpItem.data, pOpItem.Key
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Unknown order info type"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub processOrderPlexSummary(ByVal ops As OrderPlexSummary)
Const ProcName As String = "processOrderPlexSummary"
Dim opr As OrderPlexRecoverer


On Error GoTo Err

If mStartRecoveryFrom = 0 Or _
    Int(ops.CreationTime) < mStartRecoveryFrom _
    Then mStartRecoveryFrom = Int(ops.CreationTime)

On Error Resume Next
Set opr = mOrderPlexRecoverers.Item(ops.Key)
On Error GoTo Err

If opr Is Nothing Then
    Set opr = CreateOrderPlexRecoverer(ops)
    If opr Is Nothing Then Exit Sub
    opr.UpdateOrderPlex ops
    updateOrderIdToOrderPlexRecovererMap opr
Else
    opr.UpdateOrderPlex ops
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function processOrderSummary(ByVal os As OrderSummary) As Boolean
Const ProcName As String = "processOrderSummary"
Dim opr As OrderPlexRecoverer


On Error GoTo Err

On Error Resume Next
Set opr = mOrderIdToOrderPlexRecoverer.Item(os.Id)
On Error GoTo Err

If opr Is Nothing Then
    ' this can happen because order cancellations for stop/target orders
    ' may be received after the execution for the target/stop order
    ' have been received. If the target/stop order made the position
    ' flat, then a 'DELETE' item would have been recorded, but only those
    ' items recorded before the delete are actually delete.
    Exit Function
End If

mOrderIdMap.Add os.Id, os.BrokerId

opr.UpdateOrder os

processOrderSummary = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub processPositionManagerSummary( _
                ByVal pms As PositionManagerSummary, _
                ByRef pKey As String)
Const ProcName As String = "processPositionManagerSummary"
Dim pm As PositionManager


On Error GoTo Err

Set pm = gGetPositionManager(pKey)
If pm Is Nothing Then
    Set pm = gCreatePositionManager(pKey, gTB.WorkSpaces.Item(pms.WorkspaceName))
    pm.StartRecovery
End If

pm.SyncToSummary pms

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupDefaultOrderPersistenceSP()
Const ProcName As String = "setupDefaultOrderPersistenceSP"

On Error GoTo Err

gTB.ServiceProviders.Add ProjectName & ".OrderPersistenceSP", _
                        True, _
                        "RecoveryFilePath=" & ApplicationSettingsFolder

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub updateOrderIdToOrderPlexRecovererMap( _
                ByVal opr As OrderPlexRecoverer)
Const ProcName As String = "updateOrderIdToOrderPlexRecovererMap"

On Error GoTo Err

If Not opr.orderPlex.CloseoutOrder Is Nothing Then mOrderIdToOrderPlexRecoverer.Add opr, opr.orderPlex.CloseoutOrder.Id
If Not opr.orderPlex.EntryOrder Is Nothing Then mOrderIdToOrderPlexRecoverer.Add opr, opr.orderPlex.EntryOrder.Id
If Not opr.orderPlex.Stoporder Is Nothing Then mOrderIdToOrderPlexRecoverer.Add opr, opr.orderPlex.Stoporder.Id
If Not opr.orderPlex.TargetOrder Is Nothing Then mOrderIdToOrderPlexRecoverer.Add opr, opr.orderPlex.TargetOrder.Id

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


