VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockContractFetchTask"
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

Implements ITask

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

Private Const ModuleName                            As String = "MockContractFetchTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext
Private mFutureBuilder                              As New FutureBuilder

Private mContractSpecifier                          As IContractSpecifier
Private mContracts                                  As IContracts

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mListener                                   As IContractFetchListener

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

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
    mTaskContext.Finish Empty, False
    If Not mListener Is Nothing Then mListener.FetchCompleted mTaskContext.Data
End If

Exit Sub

Err:
HandleUnexpectedError ProcName, "OrderUtilsTests", ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
If Not mListener Is Nothing Then mListener.FetchCompleted mTaskContext.Data
mTaskContext.Finish Empty, False
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractsFuture() As IFuture
Set ContractsFuture = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pContracts As IContracts, _
                ByVal pListener As IContractFetchListener)
Set mContractSpecifier = pContractSpecifier
Set mContracts = pContracts
Set mListener = pListener
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





