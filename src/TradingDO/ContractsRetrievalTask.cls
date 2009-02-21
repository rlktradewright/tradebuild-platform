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

Private Const ProjectName                   As String = "TBInfoBase26"
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

Private Sub Task_cancel()
If Not mFetchInstrumentsTC Is Nothing Then mFetchInstrumentsTC.cancelTask
End Sub

Private Sub Task_run()
Dim instrument As BusinessDataObject

Dim failpoint As Long
On Error GoTo Err

If mInstrEnumerator Is Nothing Then
    Set mFetchInstrumentsTC = mDB.InstrumentFactory.queryObjectsBySpecAsync( _
                                                mBuilder.Contracts.ContractSpecifier, _
                                                , _
                                                "LoadBySpecifier/" & mBuilder.Contracts.ContractSpecifier.ToString & "(" & gGetSequenceNumber & ")")
    mTaskContext.suspend -1
Else
    If Not mInstrEnumerator.MoveNext Then
        mTaskContext.finish mBuilder.Contracts, False
    Else
        Set instrument = mInstrEnumerator.current
        mBuilder.AddContract gContractFromInstrument(instrument)
    End If
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Task_run" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
If errNumber = 3265 Then
    ' This error is 'Item cannot be found in the collection corresponding
    ' to the requested name or ordinal'. It occurs the first time this object
    ' runs after the database has not been used for some time (hours?). The
    ' instrument objects are correctly constructed, but when the first one is
    ' accessed (in gContractFromInstrument) the recordset has mysteriously been closed
    '
    ' So just try it again!
    
    gLogger.Log LogLevelNormal, "ContractsRetrievalTask restarted after receiving the following error:" & vbCrLf & _
                                "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
    Set mInstrEnumerator = Nothing
    Exit Sub
End If
    
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Property Let Task_TaskContext( _
                ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.name
End Property

'@================================================================================
' mFetchInstrumentsTC Event Handlers
'@================================================================================

Private Sub mFetchInstrumentsTC_Completed( _
                ev As TWUtilities30.TaskCompletionEvent)
Dim instruments As DataObjects

If ev.cancelled Then
    mTaskContext.finish Empty, True
ElseIf ev.errorNumber <> 0 Then
    mTaskContext.Error ev.errorNumber, ev.errorMessage, ev.errorSource
Else
    Set instruments = ev.result
    Set mInstrEnumerator = instruments.Enumerator
    
    mTaskContext.continue
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal db As TradingDB, _
                ByVal specifier As ContractSpecifier)
Set mDB = db
Set mBuilder = CreateContractsBuilder(specifier)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




