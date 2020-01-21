VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GroupResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "GroupResources"

'@================================================================================
' Member variables
'@================================================================================

Private mGroupName                                  As String

Private mContractProcessors                         As New EnumerableCollection

Private mCurrentContractProcessor                   As ContractProcessor

Private mOrderPlacers                               As New EnumerableCollection

Private mContractStore                              As IContractStore

Private mOrderManager                               As OrderManager
Private mScopeName                                  As String
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory

Private mMarketDataManager                          As IMarketDataManager

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

Dim lNewContractProcessor As ContractProcessor
Set lNewContractProcessor = ev.ContinuationData

If ev.Future.IsFaulted Then
    mContractProcessors.Remove lNewContractProcessor.Id
ElseIf ev.Future.IsAvailable Then
    Dim lContract As IContract: Set lContract = ev.Future.Value
    
    Dim lExistingContractProcessor As ContractProcessor
    
    Dim lContractProcessorName As String
    lContractProcessorName = gGenerateContractProcessorName(mGroupName, lContract.Specifier)
    If mContractProcessors.TryItem(lContractProcessorName, lExistingContractProcessor) Then
        ' we've already got a ContractProcessor for this contract, so move any
        ' pending orders from this new one to that one, and delete the new one
        lExistingContractProcessor.AddBracketOrderSpecs lNewContractProcessor
        If lNewContractProcessor Is mCurrentContractProcessor Then
            Set mCurrentContractProcessor = lExistingContractProcessor
        End If
        mContractProcessors.Remove lNewContractProcessor.Id
        
    Else
        mContractProcessors.Remove lNewContractProcessor.Id
        mContractProcessors.Add lNewContractProcessor, lContractProcessorName
        lNewContractProcessor.Activate lContractProcessorName, _
                                        lContract, _
                                        mMarketDataManager, _
                                        mOrderManager, _
                                        mScopeName, _
                                        mGroupName, _
                                        mOrderSubmitterFactory
    
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractProcessors() As EnumerableCollection
Set ContractProcessors = mContractProcessors
End Property

Friend Property Let CurrentContractProcessor(ByVal Value As ContractProcessor)
Set mCurrentContractProcessor = Value
End Property

Friend Property Get CurrentContractProcessor() As ContractProcessor
Set CurrentContractProcessor = mCurrentContractProcessor
End Property

Friend Property Get GroupName() As String
GroupName = mGroupName
End Property

Friend Property Get OrderPlacers() As EnumerableCollection
Set OrderPlacers = mOrderPlacers
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function AddContractProcessor( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContinuationData As BuySellCommandData, _
                ByVal pBatchOrders As Boolean, _
                ByVal pMaxExpenditure As Long, _
                ByVal pUnderlyingExchangeName As String) As ContractProcessor
Const ProcName As String = "AddContractProcessor"
On Error GoTo Err

Dim lContractProcessor As New ContractProcessor
lContractProcessor.Initialise GenerateGUIDString
mContractProcessors.Add lContractProcessor, lContractProcessor.Id
Set mCurrentContractProcessor = lContractProcessor

Dim lContractResolver As New ContractResolver
lContractResolver.Initialise pContractSpec, _
                            mContractStore, _
                            pBatchOrders, _
                            pMaxExpenditure, _
                            pUnderlyingExchangeName, _
                            mMarketDataManager
mFutureWaiter.Add lContractResolver.ContractFuture, lContractProcessor

Set AddContractProcessor = lContractProcessor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pGroupName As String, _
                ByVal pContractStore As IContractStore, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory)
mGroupName = pGroupName
Set mContractStore = pContractStore
Set mMarketDataManager = pMarketDataManager
Set mOrderManager = pOrderManager
mScopeName = pScopeName
Set mOrderSubmitterFactory = pOrderSubmitterFactory
End Sub

Friend Sub Purge()
Const ProcName As String = "Purge"
On Error GoTo Err

Set mCurrentContractProcessor = Nothing

Dim lOrderPlacer As OrderPlacer
For Each lOrderPlacer In mOrderPlacers
    lOrderPlacer.Finish
    gPlaceOrdersTask.RemoveOrderPlacer lOrderPlacer
Next
mOrderPlacers.Clear

Dim lPM As PositionManager
For Each lPM In mOrderManager.GetPositionManagersForGroup(mGroupName)
    lPM.Purge
Next

Dim lContractProcessor As ContractProcessor
For Each lContractProcessor In mContractProcessors
    lContractProcessor.Finish
Next
mContractProcessors.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




