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

Implements ChangeListener
Implements IOrderRecoveryListener
Implements IRecoverableObjectListener
Implements Task

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum RecoveryStates
    RecoveryStateStarting
    RecoveryStateExtractingInfo
    RecoveryStateCreatingOrderContexts
    RecoveryStateWaitingForOrderContexts
    RecoveryStateProcessingInfo
    RecoveryStateRecoveringOrders
End Enum

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

Private mOrderManager                               As OrderManager

Private mDataStore                                  As IOrderPersistenceDataStore
Private mMarketDataManager                          As IMarketDataManager
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory
Private mScopeName                                  As String

Private mBracketOrderRecoverers                     As EnumerableCollection

Private mOrderIdMap                                 As OrderIdMap

Private mOrderIdToBracketOrderRecoverer             As EnumerableCollection

Private mStartRecoveryFrom                          As Date

Private mExecutionReportsComplete                   As Boolean
Private mOrderReportsComplete                       As Boolean

Private mRecoveryAgent                              As IOrderRecoveryAgent

Private mProcessingPersistedData                    As Boolean

Private mTaskContext                                As TaskContext

Private mState                                      As RecoveryStates

Private mPositionManagers                           As PositionManagers

Private mNumOrderContextsWaiting                    As Long

Private mInMemoryStore                              As New InMemoryStore

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mBracketOrderRecoverers = New EnumerableCollection
Set mOrderIdMap = New OrderIdMap
Set mOrderIdToBracketOrderRecoverer = New EnumerableCollection
mState = RecoveryStateStarting
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "ChangeListener_Change"
On Error GoTo Err

Dim lOrderContext As OrderContext
Set lOrderContext = ev.Source

If ev.changeType = OrderContextChangeTypes.OrderContextReadyStateChanged And _
    lOrderContext.IsReady _
Then
    gLog "Order context ready", ProcName, ModuleName, lOrderContext.ContextsName & "/" & lOrderContext.Name
    mNumOrderContextsWaiting = mNumOrderContextsWaiting - 1
    If mNumOrderContextsWaiting = 0 And mState = RecoveryStateWaitingForOrderContexts Then
        mState = RecoveryStateProcessingInfo
        mTaskContext.Continue
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
restoreOrderPersistenceDataAndFail ProcName
End Sub

Private Sub IOrderRecoveryListener_EndOrderReports()
Const ProcName As String = "IOrderRecoveryListener_EndOrderReports"
On Error GoTo Err

mOrderReportsComplete = True
If mExecutionReportsComplete Then FinishRecovery

Exit Sub

Err:
restoreOrderPersistenceDataAndFail ProcName
End Sub

Private Sub IOrderRecoveryListener_NotifyExecutionReport(ByVal pExecutionReport As IExecutionReport)
Const ProcName As String = "IOrderRecoveryListener_NotifyExecutionReport"
On Error GoTo Err

Dim lOrderId As String
lOrderId = mOrderIdMap.GetOrderId(pExecutionReport.BrokerId)

If lOrderId = "" Then
    ' this can happen if the order recovery info has been deleted, so just
    ' ignore
    Exit Sub
End If

pExecutionReport.OrderId = lOrderId

Dim lBor As BracketOrderRecoverer
Set lBor = mOrderIdToBracketOrderRecoverer.Item(pExecutionReport.OrderId)

lBor.NotifyExecutionReport pExecutionReport

Exit Sub

Err:
restoreOrderPersistenceDataAndFail ProcName
End Sub

Private Sub IOrderRecoveryListener_NotifyOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "IOrderRecoveryListener_NotifyOrderReport"
On Error GoTo Err

pOrderReport.Id = mOrderIdMap.GetOrderId(pOrderReport.BrokerId)

Dim lBor As BracketOrderRecoverer
Set lBor = mOrderIdToBracketOrderRecoverer.Item(pOrderReport.Id)

lBor.NotifyOrderReport pOrderReport

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then
    ' this can happen if the order recovery info has been deleted, so just
    ' ignore
    Exit Sub
End If
restoreOrderPersistenceDataAndFail ProcName
End Sub

'@================================================================================
' IRecoverableObjectListener Interface Members
'@================================================================================

Private Sub IRecoverableObjectListener_NotifyRecoveryDataChanged(ev As RecoveryDataChangedEventData)
Const ProcName As String = "IRecoverableObjectListener_NotifyRecoveryDataChanged"
On Error GoTo Err

If mProcessingPersistedData Then Exit Sub

Dim lRecoverableObj As IRecoverable
Set lRecoverableObj = ev.Source

Dim lPositionMemberObj As IPositionMember
Set lPositionMemberObj = ev.Source

If mDataStore Is Nothing Then Exit Sub

Dim lOpItem As New OrderPersistenceItem
lOpItem.Data = lRecoverableObj.RecoveryData
lOpItem.ItemType = OrderInfoData
lOpItem.PositionManagerName = lPositionMemberObj.PositionManagerName
lOpItem.ProgId = lRecoverableObj.ProgId
lOpItem.Timestamp = GetTimestamp

mDataStore.SaveOrderInfo mScopeName, lOpItem.ToJSON

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
mTaskContext.Finish Empty, True
End Sub

Private Sub Task_Run()
Const ProcName As String = "Task_Run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    mTaskContext.Finish Empty, True
    Exit Sub
End If
    
Static sDataStoreEn As Enumerator
Static sInMemoryStoreEn As Enumerator

Dim OpItem As OrderPersistenceItem

If mState = RecoveryStateStarting Then
    If mRecoveryAgent Is Nothing Or mOrderSubmitterFactory Is Nothing Then
        mTaskContext.Finish mPositionManagers, False
        gLog "No order recovery undertaken - order submitter factory and/or order recovery agent is not configured", ProcName, ModuleName
        Exit Sub
    End If

    gLog "Persisted order info processing started", ProcName, ModuleName
    mProcessingPersistedData = True
    Set sDataStoreEn = mDataStore.GetEnumeratorForScope(mScopeName)
    mState = RecoveryStateExtractingInfo

ElseIf mState = RecoveryStateExtractingInfo Then
    If sDataStoreEn.MoveNext Then
        Dim lInfo As String
        lInfo = sDataStoreEn.Current
        processOrderPersistenceInfo lInfo
    Else
        mState = RecoveryStateCreatingOrderContexts
        If mInMemoryStore.StoreIsEmpty Then
            mProcessingPersistedData = False
            mTaskContext.Finish mPositionManagers, False
            gLog "Order recovery completed - nothing to recover", ProcName, ModuleName
            Exit Sub
        End If
        Set sInMemoryStoreEn = mInMemoryStore.Enumerator
    End If
ElseIf mState = RecoveryStateCreatingOrderContexts Then
    If sInMemoryStoreEn.MoveNext Then
        Set OpItem = sInMemoryStoreEn.Current
        
        If TypeOf OpItem.Data Is PositionManagerSummary Then
            ensurePositionManagerExists OpItem
        ElseIf TypeOf OpItem.Data Is BracketOrderSummary Then
            Dim lPm As PositionManager
            Set lPm = GetPositionManager(OpItem.PositionManagerName)
            
            If Not IsContractExpired(lPm.ContractFuture.Value) Then
                Dim lBracketOrderSummary As BracketOrderSummary
                Set lBracketOrderSummary = OpItem.Data
                getOrderContext lPm, lBracketOrderSummary.ContextName
            End If
        End If
    Else
        Set sInMemoryStoreEn = mInMemoryStore.Enumerator
        mDataStore.Clear mScopeName
        
        On Error GoTo ErrWithRestoreData
        
        If mNumOrderContextsWaiting <> 0 Then
            gLog "Waiting for order contexts to be ready", ProcName, ModuleName
            mState = RecoveryStateWaitingForOrderContexts
            mTaskContext.Suspend -1
        Else
            mState = RecoveryStateProcessingInfo
        End If
    End If
ElseIf mState = RecoveryStateProcessingInfo Then
    On Error GoTo ErrWithRestoreData
    
    If sInMemoryStoreEn.MoveNext Then
        Set OpItem = sInMemoryStoreEn.Current
        
        Dim opItemJSON As String
        opItemJSON = OpItem.ToJSON
        gLog "Processing persisted order info", pMsgQualifier:=opItemJSON, pModName:=ModuleName, pProcName:=ProcName, pLogLevel:=LogLevelDetail
    
        If processOrderPersistenceItem(OpItem) Then
            mDataStore.SaveOrderInfo mScopeName, opItemJSON
        End If
    Else
        mProcessingPersistedData = False
        gLog "Persisted order info processing completed", ProcName, ModuleName
        
        mState = RecoveryStateRecoveringOrders
    End If
ElseIf mState = RecoveryStateRecoveringOrders Then
    On Error GoTo ErrWithRestoreData
    
    gLog "Starting order recovery", ProcName, ModuleName
    mRecoveryAgent.StartOrderRecovery Me, mStartRecoveryFrom
    mTaskContext.Suspend -1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
Exit Sub

ErrWithRestoreData:
restoreOrderPersistenceDataAndFail ProcName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
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

'@================================================================================
' Methods
'@================================================================================

Friend Sub DeleteRecoveryInfo( _
                ByRef pPositionManagerName As String)
Const ProcName As String = "DeleteRecoveryInfo"
On Error GoTo Err

If Not mDataStore Is Nothing Then
    Dim OpItem As New OrderPersistenceItem
    OpItem.ItemType = OrderInfoDelete
    OpItem.PositionManagerName = pPositionManagerName
    OpItem.Timestamp = GetTimestamp
    mDataStore.SaveOrderInfo mScopeName, OpItem.ToJSON
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function RecoverOrders( _
                ByVal pOrderManager As OrderManager, _
                ByVal pPositionManagers As PositionManagers, _
                ByVal pScopeName As String, _
                ByVal pDataStore As IOrderPersistenceDataStore, _
                ByVal pRecoveryAgent As IOrderRecoveryAgent, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                Optional ByVal pCookie As Variant) As TaskController
Const ProcName As String = "RecoverOrders"
On Error GoTo Err

Static sCalled As Boolean

Assert Not sCalled, "RecoverOrders has already been called in this session for this scope, and cannot not be called again"
AssertArgument Not pPositionManagers Is Nothing, "pPositionManagers Is Nothing"
AssertArgument Not pDataStore Is Nothing, "pDataStore Is Nothing"
AssertArgument Not pMarketDataManager Is Nothing, "pMarketDataManager Is Nothing"

sCalled = True

Set mOrderManager = pOrderManager
Set mPositionManagers = pPositionManagers
mScopeName = pScopeName
Set mDataStore = pDataStore
Set mMarketDataManager = pMarketDataManager
Set mOrderSubmitterFactory = pOrderSubmitterFactory
Set mRecoveryAgent = pRecoveryAgent

Set RecoverOrders = StartTask(Me, PriorityNormal, , pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createBracketOrderRecoverer( _
                ByVal pPositionManager As PositionManager, _
                ByVal pContextName As String, _
                ByVal pKey As String) As BracketOrderRecoverer
Const ProcName As String = "createBracketOrderRecoverer"
On Error GoTo Err

Dim lBor As New BracketOrderRecoverer
lBor.Initialise pPositionManager, getOrderContext(pPositionManager, pContextName), pKey

Set createBracketOrderRecoverer = lBor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ensurePositionManagerExists(ByVal pOpItem As OrderPersistenceItem)
Const ProcName As String = "ensurePositionManagerExists"
On Error GoTo Err

If mPositionManagers.Contains(pOpItem.PositionManagerName) Then Exit Sub

Dim lSummary As PositionManagerSummary
Set lSummary = pOpItem.Data

Dim lPm As New PositionManager
lPm.PreInitialise mOrderManager

Dim lRecoverable As IRecoverable
Set lRecoverable = lPm
lRecoverable.LoadRecoveryData lSummary

Dim lContract As IContract
Set lContract = lPm.ContractFuture.Value

mPositionManagers.Add lPm

Dim lDataSource As MarketDataSource
If IsContractExpired(lContract) Then
    gLog "Discarding order recovery information for expired contract", ProcName, ModuleName, lContract.Specifier.ToString
    Exit Sub
End If
    
Set lDataSource = mMarketDataManager.CreateMarketDataSource(lPm.ContractFuture, False)
lDataSource.StartMarketData

lPm.AddRecoverableObjectListener Me
lPm.CompleteRecoveryInitialisation mOrderSubmitterFactory.CreateOrderSubmitter(lPm), lDataSource, Me

lPm.StartRecovery

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FinishRecovery()
Const ProcName As String = "finishRecovery"
On Error GoTo Err

Static sFinished As Boolean

If sFinished Then Exit Sub
sFinished = True

mRecoveryAgent.FinishOrderRecovery
Set mRecoveryAgent = Nothing

Dim lBor As BracketOrderRecoverer
For Each lBor In mBracketOrderRecoverers
    lBor.FinishRecovery
Next
Set mBracketOrderRecoverers = Nothing

Dim lPMsToRemove As New EnumerableCollection

Dim en As Enumerator
Set en = mPositionManagers.Enumerator
Do While en.MoveNext
    Dim pm As PositionManager
    Set pm = en.Current
    
    If IsContractExpired(pm.ContractFuture.Value) Then
        lPMsToRemove.Add pm
    Else
        Dim lRecoverable As IRecoverable
        Set lRecoverable = pm
        lRecoverable.FinishRecovery
        
        If pm.IsFinished Then lPMsToRemove.Add pm
    End If
Loop

Dim lPm As PositionManager
For Each lPm In lPMsToRemove
    mPositionManagers.Remove lPm
Next

mTaskContext.Finish mPositionManagers, False

gLog "Order recovery completed", ProcName, ModuleName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getOrderContext( _
                ByVal pPositionManager As PositionManager, _
                ByVal pContextName As String) As OrderContext
Const ProcName As String = "getOrderContext"
On Error GoTo Err

Dim lOrderContext As OrderContext

If pPositionManager.OrderContexts.Contains(pContextName) Then
    gLog "Order context already exists", ProcName, ModuleName, pPositionManager.Name & "/" & pContextName
    Set lOrderContext = pPositionManager.OrderContexts(pContextName)
Else
    gLog "Creating order context", ProcName, ModuleName, pPositionManager.Name & "/" & pContextName
    Set lOrderContext = pPositionManager.OrderContexts.Add(pContextName)
    lOrderContext.AddChangeListener Me
    mNumOrderContextsWaiting = mNumOrderContextsWaiting + 1
    gLog "Order contexts awaiting ready", ProcName, ModuleName, mNumOrderContextsWaiting
End If


Set getOrderContext = lOrderContext

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetPositionManager( _
                ByVal pName As String) As PositionManager
Const ProcName As String = "GetPositionManager"
On Error GoTo Err

Assert mPositionManagers.Contains(pName), "PositionManager " & pName & " not recovered"

Dim lPm As PositionManager
Set lPm = mPositionManagers.Item(pName)

Set GetPositionManager = lPm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub mapOrderIdToBracketOrderRecoverer(ByVal pOrder As IOrder, ByVal pBor As BracketOrderRecoverer)
Const ProcName As String = "mapOrderIdToBracketOrderRecoverer"
On Error GoTo Err

If pOrder Is Nothing Then Exit Sub
If mOrderIdToBracketOrderRecoverer.Contains(pOrder.Id) Then Exit Sub
mOrderIdToBracketOrderRecoverer.Add pBor, pOrder.Id

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mapOrderIdsToBracketOrderRecoverer( _
                ByVal pBor As BracketOrderRecoverer)
Const ProcName As String = "mapOrderIdsToBracketOrderRecoverer"
On Error GoTo Err

mapOrderIdToBracketOrderRecoverer pBor.BracketOrder.CloseoutOrder, pBor
mapOrderIdToBracketOrderRecoverer pBor.BracketOrder.EntryOrder, pBor
mapOrderIdToBracketOrderRecoverer pBor.BracketOrder.StopLossOrder, pBor
mapOrderIdToBracketOrderRecoverer pBor.BracketOrder.TargetOrder, pBor

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processBracketOrder(ByVal pOpi As OrderPersistenceItem)
Const ProcName As String = "processBracketOrder"
On Error GoTo Err

Dim lSummary As BracketOrderSummary
Set lSummary = pOpi.Data

If mStartRecoveryFrom = 0 Or _
    Int(lSummary.CreationTime) < mStartRecoveryFrom _
    Then mStartRecoveryFrom = Int(lSummary.CreationTime)

Dim lBor As BracketOrderRecoverer
If mBracketOrderRecoverers.Contains(lSummary.Key) Then
    Set lBor = mBracketOrderRecoverers.Item(lSummary.Key)
Else
    Dim lPm As PositionManager
    Set lPm = GetPositionManager(pOpi.PositionManagerName)
    Set lBor = createBracketOrderRecoverer(lPm, lSummary.ContextName, lSummary.Key)
    mBracketOrderRecoverers.Add lBor, lSummary.Key
End If

lBor.UpdateBracketOrder lSummary
mapOrderIdsToBracketOrderRecoverer lBor

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processExecutionReport(ByVal ex As Execution)
Const ProcName As String = "processExecutionReport"
On Error GoTo Err

Dim lBor As BracketOrderRecoverer
Set lBor = mOrderIdToBracketOrderRecoverer.Item(ex.OrderId)

lBor.UpdateExecution ex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processOrder(ByVal pOpi As OrderPersistenceItem) As Boolean
Const ProcName As String = "processOrder"
On Error GoTo Err

Dim lSummary As OrderSummary
Set lSummary = pOpi.Data

If Not mOrderIdToBracketOrderRecoverer.Contains(lSummary.Id) Then
    ' this can happen because order cancellations for stop/target orders
    ' may be received after the execution for the target/stop order
    ' have been received. If the target/stop order made the position
    ' flat, then a 'DELETE' item would have been recorded, but only those
    ' items recorded before the delete are actually deleted.
    processOrder = False
    Exit Function
End If

Dim lBor As BracketOrderRecoverer
Set lBor = mOrderIdToBracketOrderRecoverer.Item(lSummary.Id)
mOrderIdMap.Add lSummary.Id, lSummary.BrokerId

lBor.updateOrder lSummary

processOrder = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processOrderPersistenceInfo( _
                orderInfo As String)
Const ProcName As String = "processOrderPersistenceInfo"
On Error GoTo Err

Dim OpItem As OrderPersistenceItem
ParseJSON orderInfo, OpItem

If OpItem.ItemType = OrderInfoData Then
    mInMemoryStore.Add OpItem
ElseIf OpItem.ItemType = OrderInfoDelete Then
    mInMemoryStore.Delete OpItem
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processOrderPersistenceItem( _
                ByVal pOpItem As OrderPersistenceItem) As Boolean
Const ProcName As String = "processOrderPersistenceItem"
On Error GoTo Err

Dim lPm As PositionManager
Set lPm = GetPositionManager(pOpItem.PositionManagerName)

If IsContractExpired(lPm.ContractFuture.Value) Then
    processOrderPersistenceItem = False
    Exit Function
End If

processOrderPersistenceItem = True

If TypeOf pOpItem.Data Is BracketOrderSummary Then
    processBracketOrder pOpItem
ElseIf TypeOf pOpItem.Data Is OrderSummary Then
    If Not processOrder(pOpItem) Then processOrderPersistenceItem = False
ElseIf TypeOf pOpItem.Data Is Execution Then
    processExecutionReport pOpItem.Data
ElseIf TypeOf pOpItem.Data Is PositionManagerSummary Then
    processPositionManager pOpItem
Else
    Assert False, "Unknown order info type"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processPositionManager( _
                ByVal pOpi As OrderPersistenceItem)
Const ProcName As String = "processPositionManager"
On Error GoTo Err

Dim lSummary As PositionManagerSummary
Set lSummary = pOpi.Data

Dim lRecoverable As IRecoverable
Set lRecoverable = GetPositionManager(pOpi.PositionManagerName)

lRecoverable.LoadRecoveryData lSummary

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub restoreOrderPersistenceDataAndFail(ByVal pProcName As String)
Dim lErrNum As Long: lErrNum = Err.Number
Dim lErrMsg As String: lErrMsg = Err.Description
Dim lErrSource As String: lErrSource = Err.Source

Assert lErrNum <> 0, "Not called in error context"

Static sRestored As Boolean

If sRestored Then HandleUnexpectedError pProcName, ProjectName, ModuleName, , , , lErrNum, lErrMsg, lErrSource

sRestored = True

mDataStore.Clear mScopeName

Dim en As Enumerator
Set en = mInMemoryStore.Enumerator

Do While en.MoveNext
    Dim lOpItem As OrderPersistenceItem
    Set lOpItem = en.Current
    mDataStore.SaveOrderInfo mScopeName, lOpItem.ToJSON
Loop

HandleUnexpectedError pProcName, ProjectName, ModuleName, , , , lErrNum, lErrMsg, lErrSource
End Sub


