VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractRequester"
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ContractRequester"

'@================================================================================
' Member variables
'@================================================================================

Private mContractInfoServiceProvider                As ContractInfoServiceProvider
Private mCommonServiceConsumer                      As ICommonServiceConsumer
Private mHandle                                     As Long
Private mContractInfoServiceConsumer                As IContractInfoServiceConsumer

Private mName                                       As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'================================================================================
' TaskCompletionListener Interface Members
'================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

If ev.ErrorMessage <> 0 Then
    mContractInfoServiceConsumer.notifyEvent StandardSPEventCodes.CIContractSpecifierInvalid, _
                                        ev.ErrorMessage, _
                                        mHandle
    checkIfFinished
ElseIf ev.Cancelled Then
    ' nothing to do
Else
    mContractInfoServiceConsumer.LoadComplete mHandle
    checkIfFinished
    
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

Set mContractInfoServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mContractInfoServiceConsumer = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractInfoServiceProvider As ContractInfoServiceProvider, _
                ByVal CommonServiceConsumer As ICommonServiceConsumer, _
                ByVal ContractInfoServiceConsumer As IContractInfoServiceConsumer, _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractsBuilder As ContractsBuilder, _
                ByVal Name As String, _
                ByVal Handle As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

mName = Name
mHandle = Handle

Set mContractInfoServiceProvider = pContractInfoServiceProvider
Set mCommonServiceConsumer = CommonServiceConsumer
Set mContractInfoServiceConsumer = ContractInfoServiceConsumer

requestContracts pContractRequester, pContractsBuilder

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkIfFinished()
Const ProcName As String = "checkIfFinished"
On Error GoTo Err

mContractInfoServiceProvider.contractRequesterFinished Me
Finish

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.number
Dim errSource As String: errSource = Err.source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.number, Err.source, Err.Description, mHandle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub requestContracts( _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractsBuilder As ContractsBuilder)
Const ProcName As String = "requestContracts"
On Error GoTo Err

Dim lTask As New ContractDetailsRequestTask

lTask.Initialise pContractRequester, pContractsBuilder
StartTask(lTask, PriorityLow).AddTaskCompletionListener Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

