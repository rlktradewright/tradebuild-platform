VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractsRetrievalTask"
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


Private Const ModuleName                    As String = "ContractsRetrievalTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext

Private mDB                                 As TradingDB

Private mBuilder                            As ContractsBuilder

Private WithEvents mFetchInstrumentsTC      As TaskController
Attribute mFetchInstrumentsTC.VB_VarHelpID = -1

Private mInstrEnumerator                    As Enumerator

Private mHandle                             As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "ContractsRetrievalTask terminated"
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
Const ProcName As String = "Task_Cancel"
Dim failpoint As String
On Error GoTo Err

If Not mFetchInstrumentsTC Is Nothing Then mFetchInstrumentsTC.cancelTask

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Task_run()
Const ProcName As String = "Task_run"
Dim instrument As BusinessDataObject


On Error GoTo Err

If mInstrEnumerator Is Nothing Then
    Set mFetchInstrumentsTC = mDB.InstrumentFactory.QueryObjectsBySpecAsync( _
                                                mBuilder.Contracts.ContractSpecifier, _
                                                , _
                                                "LoadBySpecifier/" & mBuilder.Contracts.ContractSpecifier.ToString & "(" & gGetSequenceNumber & ")")
    mTaskContext.Suspend -1
Else
    If Not mInstrEnumerator.MoveNext Then
        mTaskContext.Finish mBuilder.Contracts, False
    Else
        Set instrument = mInstrEnumerator.current
        mBuilder.AddContract gContractFromInstrument(instrument)
    End If
End If

Exit Sub

Err:
If Err.Number = 3265 Then
    ' This error is 'Item cannot be found in the collection corresponding
    ' to the requested Name or ordinal'. It occurs the first time this object
    ' runs after the database has not been used for some time (hours?). The
    ' instrument objects are correctly constructed, but when the first one is
    ' accessed (in gContractFromInstrument) the recordset has mysteriously been closed
    '
    ' So just try it again!
    
    logMessage pMsg:="ContractsRetrievalTask restarted after receiving the following error", _
                pMsgQualifier:="Error " & Err.Number & ": " & Err.Description, _
                pProcName:="Task_run"
    Set mInstrEnumerator = Nothing
    Exit Sub
End If
    
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Let Task_TaskContext( _
                ByVal value As TWUtilities30.TaskContext)
Const ProcName As String = "Task_TaskContext"
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.Name
End Property

'@================================================================================
' mFetchInstrumentsTC Event Handlers
'@================================================================================

Private Sub mFetchInstrumentsTC_Completed( _
                ev As TWUtilities30.TaskCompletionEvent)
Const ProcName As String = "mFetchInstrumentsTC_Completed"
Dim instruments As DataObjects


On Error GoTo Err

If ev.cancelled Then
    mTaskContext.Finish Empty, True
ElseIf ev.errorNumber <> 0 Then
    mTaskContext.Error ev.errorNumber, ev.errorMessage, ev.errorSource
Else
    Set instruments = ev.result
    Set mInstrEnumerator = instruments.Enumerator
    
    mTaskContext.Continue
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal specifier As ContractSpecifier)
Const ProcName As String = "Initialise"
Set mDB = db
Set mBuilder = CreateContractsBuilder(specifier)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "logMessage"
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub





