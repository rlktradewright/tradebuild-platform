VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractLoadTask"
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


Private Const ModuleName                    As String = "ContractLoadTask"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec                       As IContractSpecifier
Private mContracts                          As IContracts
Attribute mContracts.VB_VarHelpID = -1

Private mTaskContext                        As TaskContext

Private mUsedSecondarySP                    As Boolean

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

gLogLogger.Log pMsg:="Loading contracts", pMsgQualifier:=mContractSpec.ToString, pProcName:=ProcName, pModName:=ModuleName

Dim lContractInfoSP As IContractInfoServiceProvider

Set lContractInfoSP = gTB.ServiceProviders.GetServiceProvider(ContractInfo)
If lContractInfoSP Is Nothing Then
    Set lContractInfoSP = gTB.ServiceProviders.GetServiceProvider(SecondaryContractInfo)
    If Not lContractInfoSP Is Nothing Then mUsedSecondarySP = True
End If
If lContractInfoSP Is Nothing Then
    mTaskContext.Notify ApiNotifyCodes.ApiNotifyTryLater, _
                        "No contract info service provider has been configured"
    mTaskContext.Finish mContracts, False
End If

Dim lContractStore As IContractStore
Set lContractStore = lContractInfoSP.GetContractStore
lContractStore.FetchContractDetails(mContractSpec).AddTaskCompletionListener Me
mTaskContext.suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Const ProcName As String = "Task_TaskContext"
On Error GoTo Err

Set mTaskContext = Value
Set mContractSpec = mTaskContext.Data

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get Task_TaskName() As String
Const ProcName As String = "Task_TaskName"
On Error GoTo Err

Task_TaskName = mTaskContext.name

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

If ev.errorNumber <> 0 Then
    mTaskContext.Error ev.errorNumber, ev.errorMessage, ev.source
ElseIf ev.Cancelled Then
    mTaskContext.Finish Nothing, True
Else
    Set mContracts = ev.Result

    If mContracts.Count = 0 And Not mUsedSecondarySP Then
            
        If trySecondaryContractSP Then Exit Sub
    End If
    
    gLogLogger.Log "Loaded " & mContracts.Count & " contracts: " & mContractSpec.ToString, ProcName, ModuleName
    mTaskContext.Finish mContracts, False
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

'@================================================================================
' Helper Functions
'@================================================================================

Private Function trySecondaryContractSP() As Boolean
Const ProcName As String = "trySecondaryContractSP"
On Error GoTo Err

Dim lContractInfoSP As IContractInfoServiceProvider
Set lContractInfoSP = gTB.ServiceProviders.GetServiceProvider(SecondaryContractInfo)
mUsedSecondarySP = True
If contractInfoSP Is Nothing Then
    trySecondaryContractSP = False
Else
    Dim lContractStore As IContractStore
    Set lContractStore = lContractInfoSP.GetContractStore
    lContractStore.FetchContractDetails(mContractSpec).AddTaskCompletionListener Me
    trySecondaryContractSP = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

