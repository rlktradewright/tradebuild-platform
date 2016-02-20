VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockContractStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IContractStore
Implements ITask

Private mTaskContext As TaskContext
Private mContractsBuilder As ContractsBuilder
Private mFutureBuilder As FutureBuilder
Private mListener As IContractFetchListener

'Private Function IContractStore_FetchContract(ByVal pContractSpecifier As IContractSpecifier, Optional ByVal pCookie As Variant) As IFuture
'mSingleContractRequested = True
'Set mContractsBuilder = New ContractsBuilder
'mContractsBuilder.Add gCreateContract(pContractSpecifier)
'Set mFutureBuilder = New FutureBuilder
'Set IContractStore_FetchContract = mFutureBuilder.Future
'StartTask Me, PriorityNormal, , , pCookie
'End Function

Private Function IContractStore_FetchContracts(ByVal pContractSpecifier As IContractSpecifier, Optional ByVal pListener As IContractFetchListener, Optional ByVal pCookie As Variant) As IFuture
Set mContractsBuilder = New ContractsBuilder
mContractsBuilder.Add gCreateContract(pContractSpecifier)
Set mListener = pListener
Set mFutureBuilder = New FutureBuilder
Set IContractStore_FetchContracts = mFutureBuilder.Future
StartTask Me, PriorityNormal, , , pCookie
End Function

Private Sub IContractStore_Finish()

End Sub

Private Function IContractStore_StoreContract(ByVal pContract As IContract, Optional ByVal pCookie As Variant) As TaskController
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not supported"
End Function

Private Function IContractStore_Supports(ByVal pCapabilities As ContractStoreCapabilities) As Boolean
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not supported"
End Function

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
mFutureBuilder.Value = mContractsBuilder.Contracts
mFutureBuilder.Cookie = mTaskContext.Data
mFutureBuilder.Complete
If Not mListener Is Nothing Then mListener.FetchCompleted mTaskContext.Data
mTaskContext.Finish Empty, False
End Sub

Private Property Let ITask_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get ITask_TaskName() As String

End Property
