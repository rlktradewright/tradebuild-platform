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

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mOrderPlexRecoverers = New Collection
Set mOrderIdMap = New OrderIdMap
Set mOrderIdToOrderPlexRecoverer = New Collection
End Sub

'@================================================================================
' IOrderPersistenceServiceConsumer Interface Members
'@================================================================================

Private Sub IOrderPersistenceServiceConsumer_PersistedOrderInfo( _
                ByVal pTimestamp As Date, _
                pKey As String, _
                orderInfo As String)

Const ProcName As String = "IOrderPersistenceServiceConsumer_PersistedOrderInfo"
Dim failpoint As String
On Error GoTo Err

logMessage "Processing persisted order info", pMsgQualifier:=orderInfo, pProcName:=ProcName, pLogLevel:=LogLevelDetail

mOrderPersistenceServiceProvider.SaveOrderInfo pTimestamp, pKey, orderInfo

Dim obj As Object
ParseJSON orderInfo, obj

If TypeOf obj Is OrderPlexSummary Then
    processOrderPlexSummary obj
ElseIf TypeOf obj Is OrderSummary Then
    processOrderSummary obj
ElseIf TypeOf obj Is Execution Then
    processExecution obj
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Unknown order info type"
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub IOrderPersistenceServiceConsumer_PersistedOrderInfoEnd()
Dim opr As OrderPlexRecoverer

Const ProcName As String = "IOrderPersistenceServiceConsumer_PersistedOrderInfoEnd"
Dim failpoint As String
On Error GoTo Err

logMessage "Persisted order info processing completed", pProcName:=ProcName
'For Each opr In mOrderPlexRecoverers
'    opr.PersistedOrderInfoProcessingComplete
'Next

mOrderSubmissionSrvcProvider.StartOrderRecovery Me, mStartRecoveryFrom

'!!!!!!!!!!!!!!!!!!!!!!! more here?? !!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' eg do we want to delay the next statement until we know we've had all
' the OrderReports etc from TWS?

'mOrderPersistenceServiceProvider.StartSession

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' IOrderRecoveryServiceConsumer Interface Members
'@================================================================================

Private Sub IOrderRecoveryServiceConsumer_ExecutionReport(ByVal pExecutionReport As TradeBuildSP.IExecutionReport)

Const ProcName As String = "IOrderRecoveryServiceConsumer_ExecutionReport"
Dim failpoint As String
On Error GoTo Err

Dim opr As OrderPlexRecoverer

pExecutionReport.OrderTradeBuildId = mOrderIdMap.GetTradeBuildId(pExecutionReport.OrderBrokerId)
Set opr = mOrderIdToOrderPlexRecoverer.Item(pExecutionReport.OrderTradeBuildId)

opr.NotifyExecutionReport pExecutionReport

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrInvalidProcedureCall Then
    ' this can happen if the order recovery info has been deleted, so just
    ' ignore
    Exit Sub
End If
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub IOrderRecoveryServiceConsumer_OrderReport( _
                ByVal pOrderReport As TradeBuildSP.IOrderReport)

Const ProcName As String = "IOrderRecoveryServiceConsumer_OrderReport"
Dim failpoint As Long
On Error GoTo Err

Dim opr As OrderPlexRecoverer

pOrderReport.tradeBuildId = mOrderIdMap.GetTradeBuildId(pOrderReport.BrokerId)
Set opr = mOrderIdToOrderPlexRecoverer.Item(pOrderReport.tradeBuildId)

opr.NotifyOrderReport pOrderReport

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrInvalidProcedureCall Then
    ' this can happen if the order recovery info has been deleted, so just
    ' ignore
    Exit Sub
End If
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As TWUtilities30.TaskCompletionEvent)
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

Friend Property Let IOrderSubmissionSrvcProvider( _
                ByVal value As IOrderSubmissionSrvcProvider)
Set mOrderSubmissionSrvcProvider = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub RecoverOrders(ByVal sessionName As String)
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

mOrderPersistenceServiceProvider.RequestOrderInfo sessionName
mRecovered = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Friend Sub SaveRecoveryInfo( _
                ByRef pKey As String, _
                ByRef info As String)
If Not mOrderPersistenceServiceProvider Is Nothing Then
    mOrderPersistenceServiceProvider.SaveOrderInfo GetTimestamp, pKey, info
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function CreateOrderPlexRecoverer(ByVal ops As OrderPlexSummary) As OrderPlexRecoverer
Dim ws As Workspace
Dim opr As OrderPlexRecoverer

Const ProcName As String = "CreateOrderPlexRecoverer"
Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
GLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub processExecution(ByVal ex As Execution)
Const ProcName As String = "processExecution"
Dim opr As OrderPlexRecoverer

Dim failpoint As String
On Error GoTo Err

Set opr = mOrderIdToOrderPlexRecoverer.Item(ex.OrderTradeBuildId)

opr.UpdateExecution ex

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processOrderPlexSummary(ByVal ops As OrderPlexSummary)
Const ProcName As String = "processOrderPlexSummary"
Dim opr As OrderPlexRecoverer

Dim failpoint As String
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processOrderSummary(ByVal os As OrderSummary)
Const ProcName As String = "processOrderSummary"
Dim opr As OrderPlexRecoverer

Dim failpoint As String
On Error GoTo Err

Set opr = mOrderIdToOrderPlexRecoverer.Item(os.Id)

mOrderIdMap.Add os.Id, os.BrokerId

opr.UpdateOrder os

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub setupDefaultOrderPersistenceSP()
Const ProcName As String = "setupDefaultOrderPersistenceSP"
Dim failpoint As String
On Error GoTo Err

gTB.ServiceProviders.Add ProjectName & ".OrderPersistenceSP", _
                        True, _
                        "RecoveryFilePath=" & gTB.AppSettingsFolder

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub updateOrderIdToOrderPlexRecovererMap( _
                ByVal opr As OrderPlexRecoverer)
Const ProcName As String = "updateOrderIdToOrderPlexRecovererMap"
Dim failpoint As String
On Error GoTo Err

If Not opr.orderPlex.CloseoutOrder Is Nothing Then mOrderIdToOrderPlexRecoverer.Add opr, opr.orderPlex.CloseoutOrder.Id
If Not opr.orderPlex.EntryOrder Is Nothing Then mOrderIdToOrderPlexRecoverer.Add opr, opr.orderPlex.EntryOrder.Id
If Not opr.orderPlex.stoporder Is Nothing Then mOrderIdToOrderPlexRecoverer.Add opr, opr.orderPlex.stoporder.Id
If Not opr.orderPlex.TargetOrder Is Nothing Then mOrderIdToOrderPlexRecoverer.Add opr, opr.orderPlex.TargetOrder.Id

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub


