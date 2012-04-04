VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractSelectionTask"
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

Implements TaskCompletionListener

'@================================================================================
' Events
'@================================================================================

Event Ready()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ContractSelectionTask"

'@================================================================================
' Member variables
'@================================================================================

Private mContractStore                              As IContractStore
Private mContractSelector                           As IContractSelector

Private mContractSpec                               As IContractSpecifier

Private WithEvents mFetchTaskController             As TaskController

Private mReady                                      As Boolean

Private mAllowMultipleSelection                     As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFetchTaskController Event Handlers
'@================================================================================

Private Sub mFetchTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mFetchTaskController_Completed"
On Error GoTo Err

Set mFetchTaskController = Nothing

If ev.Cancelled Then
ElseIf ev.errorNumber <> 0 Then
    mTaskContext.Error ev.errorNumber, ev.errorMessage, ev.ErrorSource
Else
    Dim lContracts As IContracts
    Set lContracts = ev.Result
    If lContracts.Count <= 1 Then
        mTaskContext.Finish lContracts, False
    Else
        mContractSelector.Initialise lContracts, mAllowMultipleSelection
        mReady = True
        RaiseEvent Ready
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get SelectedContracts() As IContracts
Const ProcName As String = "SelectedContracts"
On Error GoTo Err

Assert mReady, "Contracts not ready"
Set SelectedContracts = mContractSelector.SelectedContracts

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

If Not mFetchTaskController Is Nothing Then mFetchTaskController.CancelTask

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractStore As IContractStore, _
                ByVal pContractSelector As IContractSelector, _
                ByVal pAllowMultipleContracts As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractSpec = pContractSpec
Set mContractStore = pContractStore
Set mContractSelector = pContractSelector
mAllowMultipleSelection = pAllowMultipleContracts

Set mFetchTaskController = mContractStore.FetchContractDetails(mContractSpec)
mFetchTaskController.AddTaskCompletionListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


