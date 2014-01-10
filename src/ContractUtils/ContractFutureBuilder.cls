VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractFutureBuilder"
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

Private Const ModuleName                            As String = "ContractFutureBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskController                             As TaskController
Attribute mTaskController.VB_VarHelpID = -1
Private mReturnIContractsObject                     As Boolean

Private WithEvents mFutureBuilder                   As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureBuilder = New FutureBuilder
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

mFutureBuilder.Cookie = ev.Cookie

If ev.ErrorNumber <> 0 Then
    mFutureBuilder.Fail ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
ElseIf ev.Cancelled Then
    mFutureBuilder.Cancel
Else
    Dim lContracts As IContracts
    Set lContracts = ev.Result
    
    If mReturnIContractsObject Then
        mFutureBuilder.Value = lContracts
        mFutureBuilder.Complete
    ElseIf lContracts.Count = 0 Then
        mFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, "No such contract", ""
    ElseIf lContracts.Count > 1 Then
        mFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, "Contract not uniquely specified", ""
    Else
        mFutureBuilder.Value = lContracts.ItemAtIndex(1)
        mFutureBuilder.Complete
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureBuilder Event Handlers
'@================================================================================

Private Sub mFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mFutureBuilder_Cancelled"
On Error GoTo Err

If Not mTaskController Is Nothing Then mTaskController.CancelTask

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Future() As IFuture
Set Future = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pTaskController As TaskController, _
                ByVal pReturnIContractsObject As Boolean)
Set mTaskController = pTaskController
mTaskController.AddTaskCompletionListener Me
mReturnIContractsObject = pReturnIContractsObject
End Sub

Friend Sub InitialiseFromContract(ByVal pContract As IContract)
mFutureBuilder.Value = pContract
mFutureBuilder.Complete
End Sub

Friend Sub InitialiseFromContracts(ByVal pContracts As IContracts)
mFutureBuilder.Value = pContracts
mFutureBuilder.Complete
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

