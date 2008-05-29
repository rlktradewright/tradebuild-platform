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

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()
If Not mFetchInstrumentsTC Is Nothing Then mFetchInstrumentsTC.cancelTask
End Sub

Private Sub Task_run()
Dim instrument As BusinessDataObject

If mInstrEnumerator Is Nothing Then
    Set mFetchInstrumentsTC = mDB.InstrumentFactory.queryObjectsBySpecAsync(mBuilder.Contracts.ContractSpecifier)
    mTaskContext.sleep -1
Else
    If Not mInstrEnumerator.moveNext Then
        mTaskContext.finish mBuilder.Contracts, False
    Else
        Set instrument = mInstrEnumerator.current
        mBuilder.AddContract gContractFromInstrument(instrument)
    End If
End If
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




