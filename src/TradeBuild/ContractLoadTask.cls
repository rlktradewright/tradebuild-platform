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

Implements TradeBuildSP.IContractInfoServiceConsumer
Implements Task

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

Private mContractsBuilder                   As ContractsBuilder
Private mContracts                          As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private mTaskContext                        As TaskContext

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' DeferredAction Interface Members
'@================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_run()
Const ProcName As String = "Task_run"
On Error GoTo Err

GLogLogger.Log pMsg:="Loading contracts", pMsgQualifier:=mContractsBuilder.Contracts.ContractSpecifier.ToString, pProcName:=ProcName, pModName:=ModuleName
If Not gTB.ServiceProviders.LoadContracts(Me, mContractsBuilder) Then
    mTaskContext.notify ApiNotifyCodes.ApiNotifyTryLater, _
                        "No contract info service provider has been configured"
    mTaskContext.Finish mContracts, False
Else
    mTaskContext.suspend -1
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Dim contractSpec As ContractSpecifier

Const ProcName As String = "Task_TaskContext"
On Error GoTo Err

Set mTaskContext = RHS
Set contractSpec = mTaskContext.Data

Set mContractsBuilder = CreateContractsBuilder(contractSpec)
Set mContracts = mContractsBuilder.Contracts

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get Task_TaskName() As String
Const ProcName As String = "Task_TaskName"
On Error GoTo Err

Task_TaskName = mTaskContext.name

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' IContractInfoServiceConsumer Interface Members
'@================================================================================

Private Sub IContractInfoServiceConsumer_LoadComplete( _
                            ByVal Handle As Long)
Const ProcName As String = "IContractInfoServiceConsumer_LoadComplete"
On Error GoTo Err

If mContracts.Count = 0 Then
    ' try the secondary contract service provider if any
    If Not gTB.ServiceProviders.loadContractsFromSecondarySP(Handle, Me, mContractsBuilder) Then
        mTaskContext.Finish mContracts, False
    End If
Else
    GLogLogger.Log pMsg:="Loaded " & mContracts.Count & " contracts", pMsgQualifier:=mContractsBuilder.Contracts.ContractSpecifier.ToString, pProcName:=ProcName, pModName:=ModuleName
    mTaskContext.Finish mContracts, False
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IContractInfoServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal serviceProviderHandle As Long)

Const ProcName As String = "IContractInfoServiceConsumer_NotifyEvent"
On Error GoTo Err

Select Case eventCode
Case TradeBuildSP.StandardSPEventCodes.CIContractSpecifierInvalid
    mTaskContext.notify ApiNotifyCodes.ApiNotifyInvalidRequest, _
                        eventMessage
    mTaskContext.Finish mContracts, False
Case Else
    gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, serviceProviderHandle
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

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


