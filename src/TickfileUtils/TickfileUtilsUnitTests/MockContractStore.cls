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
Private mFutureBuilder As FutureBuilder
Private mListener As IContractFetchListener

Private mContracts As IContracts

Private mContractSpecifier As IContractSpecifier

Private WithEvents mFutureWaiter As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter

Dim lContractsBuilder As New ContractsBuilder

lContractsBuilder.Add gCreateContractFromLocalSymbol("ESM3")

lContractsBuilder.Add gCreateContractFromLocalSymbol("ZM03")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZH3")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZM3")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZU3")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZZ3")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZU4")

Set mContracts = lContractsBuilder.Contracts
End Sub

Private Function IContractStore_FetchContracts(ByVal pContractSpecifier As IContractSpecifier, Optional ByVal pListener As IContractFetchListener, Optional ByVal pCookie As Variant) As IFuture
Set mContractSpecifier = pContractSpecifier
Set mListener = pListener
Set mFutureBuilder = New FutureBuilder
Set IContractStore_FetchContracts = mFutureBuilder.Future
StartTask Me, PriorityNormal, , , pCookie
End Function

Private Function IContractStore_FetchContractsSorted(ByVal pContractSpec As ContractUtils27.IContractSpecifier, pSortkeys() As ContractUtils27.ContractSortKeyIds, Optional ByVal pSortDescending As Boolean = False, Optional ByVal pCookie As Variant) As TWUtilities40.IFuture
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not supported"
End Function

Private Function IContractStore_FetchOptionExpiries(ByVal pUnderlyingContractSpecifier As IContractSpecifier, ByVal pExchange As String, Optional ByVal pStrike As Double = 0#, Optional ByVal pCookie As Variant) As TWUtilities40.IFuture

End Function

Private Function IContractStore_FetchOptionStrikes(ByVal pUnderlyingContractSpecifier As IContractSpecifier, ByVal pExchange As String, Optional ByVal pExpiry As String, Optional ByVal pCookie As Variant) As TWUtilities40.IFuture

End Function

Private Sub IContractStore_Finish()

End Sub

Private Function IContractStore_Supports(ByVal pCapabilities As ContractStoreCapabilities) As Boolean
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not supported"
End Function

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Static sContractsBuilder As ContractsBuilder
Static en As Enumerator

If sContractsBuilder Is Nothing Then
    Set sContractsBuilder = New ContractsBuilder
    Set en = mContracts.Enumerator
End If

If en.MoveNext Then
    Dim lContract As IContract
    Set lContract = en.Current
    If ContractSpecsCompatible(mContractSpecifier, lContract.Specifier) Then
        sContractsBuilder.Add lContract
        If Not mListener Is Nothing Then mListener.NotifyContract mTaskContext.Data, lContract
    End If
Else
    mFutureWaiter.Add mFutureBuilder.Future
    mFutureBuilder.Value = sContractsBuilder.Contracts
    mFutureBuilder.Cookie = mTaskContext.Data
    mFutureBuilder.Complete
    Set en = Nothing
    Set sContractsBuilder = Nothing
    mTaskContext.Suspend -1
End If

End Sub

Private Property Let ITask_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get ITask_TaskName() As String

End Property

Private Sub mFutureWaiter_WaitCompleted(ev As TWUtilities40.FutureWaitCompletedEventData)
If Not mListener Is Nothing Then mListener.FetchCompleted mTaskContext.Data
mTaskContext.Finish Empty, False
End Sub

