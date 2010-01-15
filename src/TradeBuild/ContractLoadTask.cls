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
GLog pMsg:="Loading contracts", pMsgQualifier:=mContractsBuilder.Contracts.ContractSpecifier.ToString, pProcName:="Task_run", pProjName:=ProjectName, pModName:=ModuleName
If Not gTB.ServiceProviders.LoadContracts(Me, mContractsBuilder) Then
    mTaskContext.notify ApiNotifyCodes.ApiNotifyTryLater, _
                        "No contract info service provider has been configured"
    mTaskContext.Finish mContracts, False
Else
    mTaskContext.suspend -1
End If
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Dim contractSpec As ContractSpecifier

Set mTaskContext = RHS
Set contractSpec = mTaskContext.data

Set mContractsBuilder = CreateContractsBuilder(contractSpec)
Set mContracts = mContractsBuilder.Contracts
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.name
End Property

'@================================================================================
' IContractInfoServiceConsumer Interface Members
'@================================================================================

Private Sub IContractInfoServiceConsumer_LoadComplete( _
                            ByVal Handle As Long)
If mContracts.Count = 0 Then
    ' try the secondary contract service provider if any
    If Not gTB.ServiceProviders.loadContractsFromSecondarySP(Handle, Me, mContractsBuilder) Then
        mTaskContext.Finish mContracts, False
    End If
Else
    mTaskContext.Finish mContracts, False
End If
End Sub

Private Sub IContractInfoServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal serviceProviderHandle As Long)

Select Case eventCode
Case TradeBuildSP.StandardSPEventCodes.CIContractSpecifierInvalid
    mTaskContext.notify ApiNotifyCodes.ApiNotifyInvalidRequest, _
                        eventMessage
    mTaskContext.Finish mContracts, False
Case Else
    gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, serviceProviderHandle
End Select

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


