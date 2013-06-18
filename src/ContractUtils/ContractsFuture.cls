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

Implements IFuture

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

Private mContracts                                  As IContracts
Private mIsCancelled                                As Boolean
Private mIsCompleted                                As Boolean
Private mIsFaulted                                  As Boolean
Private mErrorNumber                                As Long
Private mErrorMessage                               As String
Private mErrorSource                                As String

Private mListeners                                  As New EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IContractsFuture Interface Members
'@================================================================================

Private Property Get IFuture_Contracts() As IContracts
Const ProcName As String = "IFuture_Contract"
On Error GoTo Err

Assert mTaskController.TaskState = TaskStates.TaskStateCompleted, "Contracts fetch task has not yet completed"

Set IFuture_Contracts = mContracts

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IFuture_Cookie() As Variant
Const ProcName As String = "IFuture_Cookie"
On Error GoTo Err

gSetVariant IFuture_Cookie, mTaskController.Cookie

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IFuture_ErrorMessage() As String
IFuture_ErrorMessage = mErrorMessage
End Property

Private Property Get IFuture_ErrorNumber() As Long
IFuture_ErrorNumber = mErrorNumber
End Property

Private Property Get IFuture_ErrorSource() As String
IFuture_ErrorSource = mErrorSource
End Property

Private Property Get IFuture_IsCancelled() As Boolean
IFuture_IsCancelled = mIsCancelled
End Property

Private Property Get IFuture_IsCompleted() As Boolean
IFuture_IsCompleted = mIsCompleted
End Property

Private Property Get IFuture_IsFaulted() As Boolean
IFuture_IsFaulted = mIsFaulted
End Property

Private Property Get IFuture_TaskController() As TaskController
Set IFuture_TaskController = mTaskController
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
    mErrorSource = ev.ErrorSource
ElseIf ev.Cancelled Then
    mIsCancelled = True
Else
    mIsCompleted = True
    Set mContracts = ev.Result
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





