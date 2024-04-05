VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FetchTicksTask"
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

Implements ITask

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


Private Const ModuleName                    As String = "FetchTicksTask"

'@================================================================================
' Member variables
'@================================================================================

Private mDB                                 As TradingDB

Private mSpecifier                          As IContractSpecifier
Private mFromDate                           As Date
Private mToDate                             As Date

Private mTaskContext                        As TaskContext

Private WithEvents mInstrumentTaskController As TaskController
Attribute mInstrumentTaskController.VB_VarHelpID = -1
                
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "FetchTicksTask terminated"
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
Const ProcName As String = "ITask_Cancel"
On Error GoTo Err

If Not mInstrumentTaskController Is Nothing Then
    mInstrumentTaskController.CancelTask
End If
mTaskContext.Finish Empty, True

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    mTaskContext.Finish Empty, True
    Exit Sub
End If

Set mInstrumentTaskController = mDB.InstrumentFactory.LoadBySpecifierAsync( _
                                                mSpecifier, _
                                                , _
                                                "LoadBySpecifier/" & mSpecifier.ToString & "(" & GetSequenceNumber & ")")
mTaskContext.Suspend -1

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_taskName() As String
ITask_taskName = mTaskContext.Name
End Property

'@================================================================================
' mInstrumentTaskController Event Handlers
'@================================================================================

Private Sub mInstrumentTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mInstrumentTaskController_Completed"
On Error GoTo Err

If ev.Cancelled Then
    mTaskContext.Finish Empty, True
ElseIf ev.ErrorNumber <> 0 Then
    mTaskContext.Error ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
ElseIf IsEmpty(ev.Result) Then
    mTaskContext.Error ErrorCodes.ErrIllegalArgumentException, _
                        "Specified object does not exist"
Else
    Dim lInstrument As instrument
    Set lInstrument = ev.Result
    mTaskContext.Finish lInstrument.CreateTickStream(mFromDate, mToDate), False
End If

Exit Sub

Err:
GTDO.NotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal Specifier As IContractSpecifier, _
                ByVal FromDate As Date, _
                ByVal ToDate As Date)
Const ProcName As String = "Initialise"
Set mDB = db
Set mSpecifier = Specifier
mFromDate = FromDate
mToDate = ToDate
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




