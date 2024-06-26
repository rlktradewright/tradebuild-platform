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

Implements IChangeListener
Implements IOrderRecoveryListener
Implements IRecoverableObjectListener
Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum RecoveryStates
    RecoveryStateStarting
    RecoveryStateExtractingInfo
    RecoveryStateCreatingPositionManagers
    RecoveryStateCreatingDataSources
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
Private mPositionManagersByRecoveryContextName      As New EnumerableCollection

Private mContractFetchesRemaining                   As Long

Private mNumOrderContextsWaiting                    As Long

Private mInMemoryStore                              As New InMemoryStore

Private mMoneyManager                               As IMoneyManager

Private mAccountDataProvider                    As IAccountDataProvider
Private mCurrencyConverter                      As ICurrencyConverter

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
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

Dim lOrderContext As OrderContext
Set lOrderContext = ev.Source

If ev.changeType = OrderContextChangeTypes.OrderContextReadyStateChanged And _
    lOrderContext.IsProviderReady _
Then
    GOrderUtils.Log "Order context ready", ProcName, ModuleName, lOrderContext.RecoveryContextName & "/" & lOrderContext.Name, LogLevelDetail
    mNumOrderContextsWaiting = mNumOrderContextsWaiting - 1
    lOrderContext.RemoveChangeListener Me
    If mNumOrderContextsWaiting = 0 And mState = RecoveryStateWaitingForOrderContexts Then
        GOrderUtils.Log "All order contexts are ready", ProcName, ModuleName
        mState = RecoveryStateProcessingInfo
        mTaskContext.Continue
    End If
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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
If mDataStore Is Nothing Then Exit Sub

Dim lRecoverableObj As IRecoverable
Set lRecoverableObj = ev.Source

Dim lOpi As New OPI
lOpi.Initialise lRecoverableObj.RecoveryContextName, _
                OrderInfoData, _
                lRecoverableObj.RecoveryData

mDataStore.SaveOrderInfo mScopeName, lOpi.ToString

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRecoverableObjectListener_WriteRecoveryDataComment(ev As RecoveryDataCommentEventData)
Const ProcName As String = "IRecoverableObjectListener_WriteRecoveryDataComment"
On Error GoTo Err

If mProcessingPersistedData Then Exit Sub
If mDataStore Is Nothing Then Exit Sub

Dim lRecoverableObj As IRecoverable
Set lRecoverableObj = ev.Source

Dim s As String
s = FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601)
s = s & "," & lRecoverableObj.RecoveryContextName
s = s & "," & OrderInfoComment
s = s & "," & ev.Comment

mDataStore.SaveOrderInfo mScopeName, s

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
mTaskContext.Finish Empty, True
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    mTaskContext.Finish Empty, True
    Exit Sub
End If
    
Static sDataStoreEn As Enumerator
Static sInMemoryStoreEn As Enumerator
Static sPositionManagersEn As Enumerator

Dim lOpItem As OPI

If mState = RecoveryStateStarting Then
    GOrderUtils.Log "Starting order recovery", ProcName, ModuleName
    If mRecoveryAgent Is Nothing Or mOrderSubmitterFactory Is Nothing Then
        mTaskContext.Finish mPositionManagers, False
        GOrderUtils.Log "No order recovery undertaken - order submitter factory and/or order recovery agent is not configured", ProcName, ModuleName
        Exit Sub
    End If

    GOrderUtils.Log "Persisted order info processing started", ProcName, ModuleName
    mProcessingPersistedData = True
    Set sDataStoreEn = mDataStore.GetEnumeratorForScope(mScopeName)
    GOrderUtils.Log "Extracting recovery data", ProcName, ModuleName
    mState = RecoveryStateExtractingInfo

ElseIf mState = RecoveryStateExtractingInfo Then
    If sDataStoreEn.MoveNext Then
        Dim lInfo As String
        lInfo = sDataStoreEn.Current
        processOrderPersistenceInfo lInfo
    Else
        If mInMemoryStore.StoreIsEmpty Then
            mProcessingPersistedData = False
            mDataStore.Clear mScopeName
            mTaskContext.Finish mPositionManagers, False
            GOrderUtils.Log "Order recovery completed - nothing to recover", ProcName, ModuleName
            Exit Sub
        End If
        Set sInMemoryStoreEn = mInMemoryStore.Enumerator
        GOrderUtils.Log "Creating position managers", ProcName, ModuleName
        mState = RecoveryStateCreatingPositionManagers
    End If
ElseIf mState = RecoveryStateCreatingPositionManagers Then
    If sInMemoryStoreEn.MoveNext Then
        Set lOpItem = sInMemoryStoreEn.Current
        
        If TypeOf lOpItem.Data Is PMS Then ensurePositionManagerExists lOpItem
    Else
        Set sPositionManagersEn = mPositionManagers.Enumerator
        mContractFetchesRemaining = mPositionManagers.Count
        GOrderUtils.Log "Creating data sources", ProcName, ModuleName
        mState = RecoveryStateCreatingDataSources
    End If
ElseIf mState = RecoveryStateCreatingDataSources Then
    If sPositionManagersEn.MoveNext Then
        Dim lPm As PositionManager: Set lPm = sPositionManagersEn.Current
        Dim lContract As IContract: Set lContract = lPm.ContractFuture.Value
        
        If IsContractExpired(lContract) Then
            GOrderUtils.Log "Discarding order recovery information for expired contract", ProcName, ModuleName, lContract.Specifier.ToString
        Else
            Dim lDataSource As MarketDataSource
            Set lDataSource = mMarketDataManager.CreateMarketDataSource(lPm.ContractFuture, False)
            lDataSource.StartMarketData
            
            lPm.CompleteRecoveryInitialisation mOrderSubmitterFactory.CreateOrderSubmitter(lDataSource), lDataSource, Me
            
            lPm.StartRecovery
        End If
    Else
        Set sInMemoryStoreEn = mInMemoryStore.Enumerator
        GOrderUtils.Log "Creating order contexts", ProcName, ModuleName
        mState = RecoveryStateCreatingOrderContexts
    End If
ElseIf mState = RecoveryStateCreatingOrderContexts Then
    If sInMemoryStoreEn.MoveNext Then
        Set lOpItem = sInMemoryStoreEn.Current
        If TypeOf lOpItem.Data Is BOS Then
            Set lPm = GetPositionManager(lOpItem)
            
            If Not IsContractExpired(lPm.ContractFuture.Value) Then
                Dim lBOS As BOS: Set lBOS = lOpItem.Data
                getOrderContext lPm, lBOS.ContextName
            End If
        End If
    Else
        Set sInMemoryStoreEn = mInMemoryStore.Enumerator
        
        On Error GoTo ErrWithRestoreData
        
        If mNumOrderContextsWaiting <> 0 Then
            GOrderUtils.Log "Waiting for order contexts to be ready", ProcName, ModuleName
            mState = RecoveryStateWaitingForOrderContexts
            mTaskContext.Suspend -1
        Else
            mDataStore.Clear mScopeName
            GOrderUtils.Log "Processing recovery data", ProcName, ModuleName
            mState = RecoveryStateProcessingInfo
        End If
    End If
ElseIf mState = RecoveryStateProcessingInfo Then
    On Error GoTo ErrWithRestoreData
    
    If sInMemoryStoreEn.MoveNext Then
        Set lOpItem = sInMemoryStoreEn.Current
        
        GOrderUtils.Log "Processing persisted order info: " & lOpItem.ToString, ProcName, ModuleName, , LogLevelDetail
    
        If processOrderPersistenceItem(lOpItem) Then
            mDataStore.SaveOrderInfo mScopeName, lOpItem.ToString
        End If
    Else
        mProcessingPersistedData = False
        GOrderUtils.Log "Persisted order info processing completed", ProcName, ModuleName
        
        GOrderUtils.Log "Recovering orders", ProcName, ModuleName
        mState = RecoveryStateRecoveringOrders
    End If
ElseIf mState = RecoveryStateRecoveringOrders Then
    On Error GoTo ErrWithRestoreData
    
    Dim en As Enumerator
    Set en = mPositionManagers.Enumerator
    Do While en.MoveNext
        Dim lRecoverable As IRecoverable
        Set lRecoverable = en.Current
        lRecoverable.FinishProcessingOrderPersistenceData
    Loop
    
    mStartRecoveryFrom = Now
    Dim lBor As BracketOrderRecoverer
    For Each lBor In mBracketOrderRecoverers
        Dim lBo As BracketOrder: Set lBo = lBor.BracketOrder
        If orderMayBeOutstanding(lBo.EntryOrder) Or _
            orderMayBeOutstanding(lBo.StopLossOrder) Or _
            orderMayBeOutstanding(lBo.TargetOrder) Or _
            orderMayBeOutstanding(lBo.CloseoutOrder) _
        Then
            If Int(lBor.UpdateTime) < mStartRecoveryFrom Then
                ' we subtract 1 second from the update time to allow
                ' for imprecision in update times
                mStartRecoveryFrom = lBor.UpdateTime - 1 / 86400
            End If
        End If
    Next

    GOrderUtils.Log "Starting order recovery", ProcName, ModuleName
    mRecoveryAgent.StartOrderRecovery Me, mStartRecoveryFrom
    mTaskContext.Suspend -1
End If

Exit Sub

Err:
mTaskContext.Error Err.Number, Err.Description, Err.Source
Exit Sub

ErrWithRestoreData:
restoreOrderPersistenceDataAndFail ProcName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ScopeName() As String
ScopeName = mScopeName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub DeleteRecoveryInfo( _
                ByRef pRecoveryContextName As String)
Const ProcName As String = "DeleteRecoveryInfo"
On Error GoTo Err

If Not mDataStore Is Nothing Then
    Dim s As String
    s = FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601)
    s = s & "," & pRecoveryContextName
    s = s & "," & OrderInfoDelete
    
    mDataStore.SaveOrderInfo mScopeName, s
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function RecoverOrders( _
                ByVal pOrderManager As OrderManager, _
                ByVal pPositionManagers As PositionManagers, _
                ByVal pScopeName As String, _
                ByVal pDataStore As IOrderPersistenceDataStore, _
                ByVal pRecoveryAgent As IOrderRecoveryAgent, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                Optional ByVal pMoneyManager As IMoneyManager, _
                Optional ByVal pAccountDataProvider As IAccountDataProvider, _
                Optional ByVal pCurrencyConverter As ICurrencyConverter, _
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
Set mMoneyManager = pMoneyManager
Set mAccountDataProvider = pAccountDataProvider
Set mCurrencyConverter = pCurrencyConverter
Set mRecoveryAgent = pRecoveryAgent

Set RecoverOrders = StartTask(Me, PriorityNormal, , pCookie)

GOrderUtils.Log "Order recovery task started", ProcName, ModuleName

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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
lBor.Initialise pPositionManager, getOrderContext(pPositionManager, pContextName), pKey, Me

Set createBracketOrderRecoverer = lBor

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ensurePositionManagerExists(ByVal pOpItem As OPI)
Const ProcName As String = "ensurePositionManagerExists"
On Error GoTo Err

Dim lSummary As PMS
Set lSummary = pOpItem.Data

Dim lPm As PositionManager
Dim lRecoverable As IRecoverable

If Not mPositionManagers.Contains(lSummary.Name) Then
    Set lPm = New PositionManager
    lPm.PreInitialise mOrderManager, _
                    mMoneyManager, _
                    mAccountDataProvider, _
                    mCurrencyConverter
    
    Set lRecoverable = lPm
    lRecoverable.LoadRecoveryData pOpItem.RCN, lSummary
    
    mPositionManagers.Add lPm
Else
    Set lPm = mPositionManagers.Item(lSummary.Name)
End If

If Not mPositionManagersByRecoveryContextName.Contains(pOpItem.RCN) Then
    mPositionManagersByRecoveryContextName.Add lPm, pOpItem.RCN
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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
    lBor.FinishDataRecovery
Next

Dim lPMsToFinish As New EnumerableCollection
Dim en As Enumerator
Set en = mPositionManagers.Enumerator
Do While en.MoveNext
    Dim lRecoverable As IRecoverable
    Set lRecoverable = en.Current
    lRecoverable.FinishRecovery

    Dim pm As PositionManager
    Set pm = en.Current
    
    Dim lContract As IContract: Set lContract = pm.ContractFuture.Value
    If IsContractExpired(pm.ContractFuture.Value) Then
        GOrderUtils.Log "Finishing position manager for expired contract: " & lContract.ToString, ProcName, ModuleName, , LogLevelDetail
        lPMsToFinish.Add pm
    ElseIf Not pm.IsActive Then
        GOrderUtils.Log "Finishing inactive position manager: " & lContract.ToString, ProcName, ModuleName, , LogLevelDetail
        lPMsToFinish.Add pm
    End If
Loop

For Each lBor In mBracketOrderRecoverers
    lBor.FinishStateRecovery
Next
Set mBracketOrderRecoverers = Nothing

Dim lPm As PositionManager
For Each lPm In lPMsToFinish
    ' note that calling Finish removes the PositionManager from the
    ' PositionManagers collection
    lPm.Finish
Next

mTaskContext.Finish mPositionManagers, False

GOrderUtils.Log "Order recovery completed", ProcName, ModuleName

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getOrderContext( _
                ByVal pPositionManager As PositionManager, _
                ByVal pContextName As String) As OrderContext
Const ProcName As String = "getOrderContext"
On Error GoTo Err

Dim lOrderContext As OrderContext

If pPositionManager.OrderContexts.Contains(pContextName) Then
    GOrderUtils.Log "Order context already exists", ProcName, ModuleName, pPositionManager.Name & "/" & pContextName, LogLevelDetail
    Set lOrderContext = pPositionManager.OrderContexts(pContextName)
Else
    GOrderUtils.Log "Creating order context", ProcName, ModuleName, pPositionManager.Name & "/" & pContextName, LogLevelDetail
    Set lOrderContext = pPositionManager.OrderContexts.Add(pContextName)
    If lOrderContext.IsProviderReady Then
        GOrderUtils.Log "Order context ready", ProcName, ModuleName, lOrderContext.RecoveryContextName & "/" & lOrderContext.Name, LogLevelDetail
    Else
        lOrderContext.AddChangeListener Me
        mNumOrderContextsWaiting = mNumOrderContextsWaiting + 1
        GOrderUtils.Log "Order contexts awaiting ready", ProcName, ModuleName, mNumOrderContextsWaiting, LogLevelDetail
    End If
End If

Set getOrderContext = lOrderContext

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetPositionManager( _
                ByVal lOpItem As OPI) As PositionManager
Const ProcName As String = "GetPositionManager"
On Error GoTo Err

Dim lPm As PositionManager
If Not mPositionManagersByRecoveryContextName.TryItem(lOpItem.RCN, lPm) Then
    GOrderUtils.Log "PositionManager " & lOpItem.RCN & " not recovered", ProcName, ModuleName, , LogLevelDetail
End If

Set GetPositionManager = lPm

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub mapOrderIdToBracketOrderRecoverer(ByVal pOrder As IOrder, ByVal pBor As BracketOrderRecoverer)
Const ProcName As String = "mapOrderIdToBracketOrderRecoverer"
On Error GoTo Err

If pOrder Is Nothing Then Exit Sub
If mOrderIdToBracketOrderRecoverer.Contains(pOrder.Id) Then Exit Sub
mOrderIdToBracketOrderRecoverer.Add pBor, pOrder.Id

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function orderMayBeOutstanding(ByVal pOrder As IOrder) As Boolean
If pOrder Is Nothing Then
    orderMayBeOutstanding = False
ElseIf pOrder.Status = OrderStatusCancelled Or _
    pOrder.Status = OrderStatusFilled Or _
    pOrder.Status = OrderStatusRejected _
Then
    orderMayBeOutstanding = False
Else
    orderMayBeOutstanding = True
End If
End Function

Private Sub processBracketOrder(ByVal pOpi As OPI)
Const ProcName As String = "processBracketOrder"
On Error GoTo Err

Dim lSummary As BOS
Set lSummary = pOpi.Data

Dim lBor As BracketOrderRecoverer
If mBracketOrderRecoverers.Contains(lSummary.Key) Then
    Set lBor = mBracketOrderRecoverers.Item(lSummary.Key)
Else
    Dim lPm As PositionManager
    Set lPm = GetPositionManager(pOpi)
    Set lBor = createBracketOrderRecoverer(lPm, lSummary.ContextName, lSummary.Key)
    mBracketOrderRecoverers.Add lBor, lSummary.Key
End If

lBor.UpdateBracketOrder pOpi.RCN, lSummary, pOpi.Timestamp
mapOrderIdsToBracketOrderRecoverer lBor

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processExecutionReport(ByVal pOpi As OPI)
Const ProcName As String = "processExecutionReport"
On Error GoTo Err

Dim lExec As Execution
Set lExec = pOpi.Data

lExec.SetClean

Dim lBor As BracketOrderRecoverer
Set lBor = mOrderIdToBracketOrderRecoverer.Item(lExec.OrderId)

lBor.UpdateExecution lExec

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processOrder(ByVal pOpi As OPI) As Boolean
Const ProcName As String = "processOrder"
On Error GoTo Err

Dim lSummary As OS
Set lSummary = pOpi.Data

If Not mOrderIdToBracketOrderRecoverer.Contains(lSummary.Id) Then
    ' this can happen because order cancellations for stop/target orders
    ' may be received after the execution for the target/stop-loss order
    ' have been received. If the target/stop-loss order made the position
    ' flat, then a 'DELETE' item would have been recorded, but only those
    ' items recorded before the delete are actually deleted.
    processOrder = False
    Exit Function
End If

Dim lBor As BracketOrderRecoverer
Set lBor = mOrderIdToBracketOrderRecoverer.Item(lSummary.Id)

mOrderIdMap.Add lSummary.Id, lSummary.BrokerId

lBor.updateOrder pOpi.RCN, lSummary

processOrder = True

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processOrderPersistenceInfo( _
                orderInfo As String)
Const ProcName As String = "processOrderPersistenceInfo"
On Error GoTo Err

Dim OpItem As New OPI: OpItem.InitialiseFromString orderInfo

If OpItem.ItemType = OrderInfoData Then
    mInMemoryStore.Add OpItem
ElseIf OpItem.ItemType = OrderInfoComment Then
    mInMemoryStore.Add OpItem
ElseIf OpItem.ItemType = OrderInfoDelete Then
    mInMemoryStore.Delete OpItem
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processOrderPersistenceItem( _
                ByVal pOpItem As OPI) As Boolean
Const ProcName As String = "processOrderPersistenceItem"
On Error GoTo Err

Dim lPm As PositionManager
Set lPm = GetPositionManager(pOpItem)
If lPm Is Nothing Then
    processOrderPersistenceItem = False
    Exit Function
End If

Debug.Print "PositionManager object key: " & GetObjectKey(lPm)

If IsContractExpired(lPm.ContractFuture.Value) Then
    processOrderPersistenceItem = False
    Exit Function
End If

processOrderPersistenceItem = True

If TypeOf pOpItem.Data Is BOS Then
    processBracketOrder pOpItem
ElseIf TypeOf pOpItem.Data Is OS Then
    If Not processOrder(pOpItem) Then processOrderPersistenceItem = False
ElseIf TypeOf pOpItem.Data Is Execution Then
    processExecutionReport pOpItem
ElseIf TypeOf pOpItem.Data Is PMS Then
    processPositionManager pOpItem
ElseIf pOpItem.ItemType = OrderInfoComment Then

Else
    Assert False, "Unknown order info type"
End If

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processPositionManager( _
                ByVal pOpi As OPI)
Const ProcName As String = "processPositionManager"
On Error GoTo Err

Dim lSummary As PMS
Set lSummary = pOpi.Data

Dim lRecoverable As IRecoverable
Set lRecoverable = GetPositionManager(pOpi)

lRecoverable.LoadRecoveryData pOpi.RCN, lSummary

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub restoreOrderPersistenceDataAndFail(ByVal pProcName As String)
Dim lErrNum As Long: lErrNum = Err.Number
Dim lErrMsg As String: lErrMsg = Err.Description
Dim lErrSource As String: lErrSource = Err.Source

Assert lErrNum <> 0, "Not called in error context"

Static sRestored As Boolean

If sRestored Then TWUtilities40.HandleUnexpectedError pProcName, ProjectName, ModuleName, , , , lErrNum, lErrMsg, lErrSource

sRestored = True

mDataStore.Clear mScopeName

Dim en As Enumerator
Set en = mInMemoryStore.Enumerator

Do While en.MoveNext
    Dim lOpItem As OPI
    Set lOpItem = en.Current
    mDataStore.SaveOrderInfo mScopeName, lOpItem.ToString
Loop

TWUtilities40.HandleUnexpectedError pProcName, ProjectName, ModuleName, , , , lErrNum, lErrMsg, lErrSource
End Sub


