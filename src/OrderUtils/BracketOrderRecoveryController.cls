VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderRecoveryCtlr"
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

Implements IOrderRecoveryListener
Implements IRecoverableObjectListener
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

Private Const ModuleName                            As String = "BracketOrderRecoveryCtlr"

Private Const OrderInfoDelete                       As String = "DELETE"
Private Const OrderInfoData                         As String = "DATA"

'@================================================================================
' Member variables
'@================================================================================

Private mDataStore                                  As IOrderPersistenceDataStore

Private mMarketDataManager                          As MarketDataManager

Private mRecovered                                  As Boolean

Private mBracketOrderRecoverers                     As Collection

Private mOrderIdMap                                 As OrderIdMap

Private mOrderIdToBracketOrderRecoverer             As Collection

Private mStartRecoveryFrom                          As Date

Private mOrderPersistenceDataGroups                 As EnumerableCollection

Private mExecutionReportsComplete                   As Boolean
Private mOrderReportsComplete                       As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mBracketOrderRecoverers = New Collection
Set mOrderIdMap = New OrderIdMap
Set mOrderIdToBracketOrderRecoverer = New Collection
Set mOrderPersistenceDataGroups = New EnumerableCollection
End Sub

'@================================================================================
' IOrderRecoveryListener Interface Members
'@================================================================================

Private Sub IOrderRecoveryListener_EndExecutionReports()
Const ProcName As String = "IOrderRecoveryListener_EndExecutionReports"
On Error GoTo Err

mExecutionReportsComplete = True
If mOrderReportsComplete Then FinishRecovery

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderRecoveryListener_EndOrderReports()
Const ProcName As String = "IOrderRecoveryListener_EndOrderReports"
On Error GoTo Err

mOrderReportsComplete = True
If mExecutionReportsComplete Then FinishRecovery

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderRecoveryListener_NotifyExecutionReport(ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "IOrderRecoveryListener_NotifyExecutionReport"
On Error GoTo Err

Dim opr As BracketOrderRecoverer
Dim tbId As String

tbId = mOrderIdMap.GetTradeBuildId(pExecutionReport.OrderBrokerId)

If tbId = "" Then
    ' this can happen if the order recovery info has been deleted, so just
    ' ignore
    Exit Sub
End If

pExecutionReport.OrderTradeBuildId = tbId
Set opr = mOrderIdToBracketOrderRecoverer.Item(pExecutionReport.OrderTradeBuildId)

opr.NotifyExecutionReport pExecutionReport

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderRecoveryListener_NotifyOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "IOrderRecoveryListener_NotifyOrderReport"
On Error GoTo Err

Dim opr As BracketOrderRecoverer

pOrderReport.TradeBuildId = mOrderIdMap.GetTradeBuildId(pOrderReport.BrokerId)
Set opr = mOrderIdToBracketOrderRecoverer.Item(pOrderReport.TradeBuildId)

opr.NotifyOrderReport pOrderReport

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then
    ' this can happen if the order recovery info has been deleted, so just
    ' ignore
    Exit Sub
End If
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' IRecoverableObjectListener Interface Members
'@================================================================================

Private Sub IRecoverableObjectListener_NotifyRecoveryDataChanged(ev As RecoveryDataChangedEventData)
Const ProcName As String = "IRecoverableObjectListener_NotifyRecoveryDataChanged"
On Error GoTo Err

SaveRecoveryInfo ev.Key, ev.ProgId, ev.Data

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As TaskCompletionEventData)
Dim opr As BracketOrderRecoverer
Dim op As BracketOrder

Set opr = mBracketOrderRecoverers.Item(ev.cookie)
Set op = opr.BracketOrder

mBracketOrderRecoverers.Remove ev.cookie
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub DeleteRecoveryInfo( _
                ByRef pKey As String)
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

If Not mDataStore Is Nothing Then
    Dim opItem As New OrderPersistenceItem
    opItem.ItemType = OrderInfoDelete
    opItem.Key = pKey
    opItem.Timestamp = GetTimestamp
    mDataStore.SaveOrderInfo opItem.ToJSON
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RecoverOrders( _
                ByVal pScopeName As String, _
                ByVal pSessionId As String, _
                ByVal pDataStore As IOrderPersistenceDataStore, _
                ByVal pRecoveryAgent As IOrderRecoveryAgent, _
                ByVal pMarketDataManager As MarketDataManager)
Const ProcName As String = "RecoverOrders"
On Error GoTo Err

Dim en As Enumerator
Dim lInfo As String

If mRecovered Then Err.Raise ErrorCodes.ErrIllegalStateException, , "RecoverOrders has already been called and may not be called again"
If pDataStore Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "pDataStore must be supplied"
If pMarketDataManager Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "pMarketDataManager must be supplied"

Set mDataStore = pDataStore
Set mMarketDataManager = pMarketDataManager

gLog "Persisted order info processing started", ModuleName, ProcName

Set en = mDataStore.GetEnumeratorForScope(pScopeName)
Do While en.MoveNext
    lInfo = en.Current
    processOrderInfo lInfo
Loop

mRecovered = True

mDataStore.StartSession pSessionId

Do While getNextOrderPersistenceItem(opItem)
    Dim opItemJSON As String
    opItemJSON = opItem.ToJSON
    gLog "Processing persisted order info", pMsgQualifier:=opItemJSON, pModName:=ModuleName, pProcName:=ProcName, pLogLevel:=LogLevelDetail

    If processOrderPersistenceItem(opItem) Then
        mDataStore.SaveOrderInfo opItemJSON
    End If
Loop

pRecoveryAgent.StartOrderRecovery Me, mStartRecoveryFrom

gLog "Persisted order info processing completed", ModuleName, ProcName

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub SaveRecoveryInfo( _
                ByVal pKey As String, _
                ByVal pProgId As String, _
                ByVal pInfo As Variant)
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

If pKey = "" Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "pKey cannot be blank"
End If

If Not mDataStore Is Nothing Then
    Dim opItem As New OrderPersistenceItem
    opItem.Data = pInfo
    opItem.ProgId = pProgId
    opItem.ItemType = OrderInfoData
    opItem.Key = pKey
    opItem.Timestamp = GetTimestamp
    mDataStore.SaveOrderInfo opItem.ToJSON
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function CreateBracketOrderRecoverer(ByVal pBos As BracketOrderSummary) As BracketOrderRecoverer
Const ProcName As String = "CreateBracketOrderRecoverer"
On Error GoTo Err

Dim lBor As BracketOrderRecoverer

Set lBor = New BracketOrderRecoverer
lBor.Initialise pBos.Key, pBos.WorkspaceName, pBos.TickerKey
StartTask(lBor, PriorityNormal, , lBor.Key).AddTaskCompletionListener Me

mBracketOrderRecoverers.Add lBor, lBor.Key

Set CreateBracketOrderRecoverer = lBor

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub FinishRecovery()
Static lFinished As Boolean

Dim opr As BracketOrderRecoverer

Const ProcName As String = "finishRecovery"
On Error GoTo Err

If lFinished Then Exit Sub
lFinished = True

For Each opr In mBracketOrderRecoverers
    opr.FinishRecovery
Next

Set mBracketOrderRecoverers = Nothing

Dim en As Enumerator
Dim pm As PositionManager
Set en = gGetPositionManagersEnumerator
Do While en.MoveNext
    Set pm = en.Current
    If Not pm.FinishRecovery Then en.Remove
Loop

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub processBracketOrderSummary(ByVal pBos As BracketOrderSummary)
Const ProcName As String = "processBracketOrderSummary"
Dim lBor As BracketOrderRecoverer


On Error GoTo Err

If mStartRecoveryFrom = 0 Or _
    Int(pBos.CreationTime) < mStartRecoveryFrom _
    Then mStartRecoveryFrom = Int(pBos.CreationTime)

On Error Resume Next
Set lBor = mBracketOrderRecoverers.Item(pBos.Key)
On Error GoTo Err

If lBor Is Nothing Then
    Set lBor = CreateBracketOrderRecoverer(pBos)
    If lBor Is Nothing Then Exit Sub
    lBor.UpdateBracketOrder pBos
    updateOrderIdToBracketOrderRecovererMap lBor
Else
    lBor.UpdateBracketOrder pBos
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processExecution(ByVal ex As Execution)
Const ProcName As String = "processExecution"
Dim opr As BracketOrderRecoverer


On Error GoTo Err

Set opr = mOrderIdToBracketOrderRecoverer.Item(ex.OrderTradeBuildId)

opr.UpdateExecution ex

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processOrderInfo( _
                orderInfo As String)

Const ProcName As String = "processOrderInfo"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function processOrderPersistenceItem( _
                ByVal pOpItem As OrderPersistenceItem) As Boolean
Const ProcName As String = "processOrderPersistenceItem"
On Error GoTo Err

processOrderPersistenceItem = True

If TypeOf pOpItem.Data Is BracketOrderSummary Then
    processBracketOrderSummary pOpItem.Data
ElseIf TypeOf pOpItem.Data Is OrderSummary Then
    If Not processOrderSummary(pOpItem.Data) Then processOrderPersistenceItem = False
ElseIf TypeOf pOpItem.Data Is Execution Then
    processExecution pOpItem.Data
ElseIf TypeOf pOpItem.Data Is PositionManagerSummary Then
    processPositionManagerSummary pOpItem.Data, pOpItem.Key
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Unknown order info type"
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function processOrderSummary(ByVal os As OrderSummary) As Boolean
Const ProcName As String = "processOrderSummary"
On Error GoTo Err

Dim opr As BracketOrderRecoverer

On Error Resume Next
Set opr = mOrderIdToBracketOrderRecoverer.Item(os.Id)
On Error GoTo Err

If opr Is Nothing Then
    ' this can happen because order cancellations for stop/target orders
    ' may be received after the execution for the target/stop order
    ' have been received. If the target/stop order made the position
    ' flat, then a 'DELETE' item would have been recorded, but only those
    ' items recorded before the delete are actually deleted.
    Exit Function
End If

mOrderIdMap.Add os.Id, os.BrokerId

opr.UpdateOrder os

processOrderSummary = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub processPositionManagerSummary( _
                ByVal pms As PositionManagerSummary, _
                ByRef pKey As String)
Const ProcName As String = "processPositionManagerSummary"
On Error GoTo Err

Dim pm As PositionManager

Set pm = gGetPositionManager(pKey)
If pm Is Nothing Then
    Set pm = gCreatePositionManager(pKey, gTB.WorkSpaces.Item(pms.WorkspaceName))
    pm.StartRecovery
End If

pm.SyncToSummary pms

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupDefaultOrderPersistenceSP()
Const ProcName As String = "setupDefaultOrderPersistenceSP"

On Error GoTo Err

gTB.ServiceProviders.Add ProjectName & ".OrderPersistenceSP", _
                        True, _
                        "RecoveryFilePath=" & ApplicationSettingsFolder

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub updateOrderIdToBracketOrderRecovererMap( _
                ByVal opr As BracketOrderRecoverer)
Const ProcName As String = "updateOrderIdToBracketOrderRecovererMap"

On Error GoTo Err

If Not opr.BracketOrder.CloseoutOrder Is Nothing Then mOrderIdToBracketOrderRecoverer.Add opr, opr.BracketOrder.CloseoutOrder.Id
If Not opr.BracketOrder.EntryOrder Is Nothing Then mOrderIdToBracketOrderRecoverer.Add opr, opr.BracketOrder.EntryOrder.Id
If Not opr.BracketOrder.Stoporder Is Nothing Then mOrderIdToBracketOrderRecoverer.Add opr, opr.BracketOrder.Stoporder.Id
If Not opr.BracketOrder.TargetOrder Is Nothing Then mOrderIdToBracketOrderRecoverer.Add opr, opr.BracketOrder.TargetOrder.Id

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


