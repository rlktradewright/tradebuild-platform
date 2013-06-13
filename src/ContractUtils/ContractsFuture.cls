VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractsFuture"
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

Implements IContractFuture

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

Private Const ModuleName                            As String = "ContractsFuture"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTaskController                  As TaskController
Attribute mTaskController.VB_VarHelpID = -1

Private mContract                                   As IContract
Private mIsCancelled                                As Boolean
Private mIsCompleted                                As Boolean
Private mIsFaulted                                  As Boolean
Private mErrorNumber                                  As Long
Private mErrorMessage                               As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IContractFuture Interface Members
'@================================================================================

Private Property Get IContractFuture_Contract() As IContract
Const ProcName As String = "IContractFuture_Contract"
On Error GoTo Err

Assert mTaskController.TaskState = TaskStates.TaskStateCompleted, "Contract fetch task has not yet completed"

Set IContractFuture_Contract = mContract

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IContractFuture_ErrorCode() As Long
IContractFuture_ErrorCode = mErrorNumber
End Property

Private Property Get IContractFuture_ErrorMessage() As String
IContractFuture_ErrorMessage = mErrorMessage
End Property

Private Property Get IContractFuture_ErrorNumber() As Long
IContractFuture_ErrorNumber = mErrorNumber
End Property

Private Property Get IContractFuture_IsCancelled() As Boolean
IContractFuture_IsCancelled = mIsCancelled
End Property

Private Property Get IContractFuture_IsCompleted() As Boolean
IContractFuture_IsCompleted = mIsCompleted
End Property

Private Property Get IContractFuture_IsFaulted() As Boolean
IContractFuture_IsFaulted = mIsFaulted
End Property

Private Property Get IContractFuture_TaskController() As TaskController
Set IContractFuture_TaskController = mTaskController
End Property

'@================================================================================
' mTaskController Event Handlers
'@================================================================================

Private Sub mTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mTaskController_Completed"
On Error GoTo Err

If ev.ErrorNumber <> 0 Then
    mIsFaulted = True
    mErrorNumber = ev.ErrorNumber
    mErrorMessage = ev.ErrorMessage
ElseIf ev.Cancelled Then
    mIsCancelled = True
Else
    Dim lContracts As IContracts
    Set lContracts = ev.Result
    
    If lContracts.Count = 0 Then
        mIsFaulted = True
        mErrorNumber = ErrorCodes.ErrIllegalArgumentException
        mErrorMessage = "No such contract"
    ElseIf lContracts.Count > 1 Then
        mIsFaulted = True
        mErrorNumber = ErrorCodes.ErrIllegalArgumentException
        mErrorMessage = "Contract not uniquely specified"
    Else
        mIsCompleted = True
        Set mContract = lContracts.Item(1)
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pTaskController As TaskController)
Set mTaskController = pTaskController
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





