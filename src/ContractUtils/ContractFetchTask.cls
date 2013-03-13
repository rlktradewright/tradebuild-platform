VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractFetchTask"
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

Implements Task
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

Private Const ModuleName                            As String = "ContractFetchTask"

'@================================================================================
' Member variables
'@================================================================================

Private mPrimaryContractStore                       As IContractStore
Private mSecondaryContractStore                     As IContractStore

Private mContractSpec                               As IContractSpecifier
Private mContracts                                  As IContracts

Private mTaskContext                                As TaskContext

Private mUsedSecondaryContractStore                 As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_run()
Const ProcName As String = "Task_run"
On Error GoTo Err

gLogger.Log pMsg:="Loading contracts", pMsgQualifier:=mContractSpec.ToString, pProcName:=ProcName, pModName:=ModuleName

Dim lContractStore As IContractStore
Set lContractStore = mPrimaryContractStore
If lContractStore Is Nothing Then
    Set lContractStore = mSecondaryContractStore
    mUsedSecondaryContractStore = True
End If

lContractStore.FetchContractDetails(mContractSpec).AddTaskCompletionListener Me
mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Const ProcName As String = "Task_TaskContext"
On Error GoTo Err

Set mTaskContext = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get Task_TaskName() As String
Const ProcName As String = "Task_TaskName"
On Error GoTo Err

Task_TaskName = mTaskContext.Name

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

If ev.ErrorNumber <> 0 Then
    mTaskContext.Error ev.ErrorNumber, ev.ErrorMessage, ev.Source
ElseIf ev.Cancelled Then
    mTaskContext.Finish Nothing, True
Else
    Set mContracts = ev.Result

    If mContracts.Count = 0 And Not mUsedSecondaryContractStore Then If trySecondaryContractSP Then Exit Sub
    
    mTaskContext.Finish mContracts, False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore)
Set mContractSpec = pContractSpec
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function trySecondaryContractSP() As Boolean
Const ProcName As String = "trySecondaryContractSP"
On Error GoTo Err

mUsedSecondaryContractStore = True
If mSecondaryContractStore Is Nothing Then
    trySecondaryContractSP = False
Else
    mSecondaryContractStore.FetchContractDetails(mContractSpec).AddTaskCompletionListener Me
    trySecondaryContractSP = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




