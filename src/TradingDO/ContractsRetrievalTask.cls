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


Private Const ModuleName                    As String = "ContractsRetrievalTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext

Private mDB                                 As TradingDB

Private mListener                           As IContractFetchListener
Private mContractSpec                       As IContractSpecifier

Private mBuilder                            As New ContractsBuilder

Private WithEvents mFetchInstrumentsTC      As TaskController
Attribute mFetchInstrumentsTC.VB_VarHelpID = -1

Private mInstrEnumerator                    As Enumerator

Private WithEvents mFutureBuilder           As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1

Private WithEvents mFutureWaiter             As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureBuilder = New FutureBuilder
Set mFutureWaiter = New FutureWaiter
End Sub

Private Sub Class_Terminate()
Debug.Print "ContractsRetrievalTask terminated"
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
Const ProcName As String = "ITask_Cancel"
On Error GoTo Err

doCancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    doCancel
    Exit Sub
End If

If mInstrEnumerator Is Nothing Then
    Set mFetchInstrumentsTC = mDB.InstrumentFactory.QueryObjectsBySpecAsync( _
                                                mContractSpec, _
                                                , _
                                                "LoadBySpecifier/" & mContractSpec.ToString & "(" & gGetSequenceNumber & ")")
    mTaskContext.Suspend -1
Else
    If Not mInstrEnumerator.MoveNext Then
        mFutureWaiter.Add mFutureBuilder.Future
        mFutureBuilder.Value = mBuilder.Contracts
        mFutureBuilder.Complete
        mTaskContext.Suspend -1
    Else
        Dim lInstrument As BusinessDataObject
        Set lInstrument = mInstrEnumerator.Current
        
        Dim lContract As IContract
        Set lContract = gContractFromInstrument(lInstrument)
        
        mBuilder.Add lContract
        If Not mListener Is Nothing Then mListener.NotifyContract mFutureBuilder.cookie, lContract
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
                pProcName:="ITask_run"
    Set mInstrEnumerator = Nothing
    Exit Sub
End If
    
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext( _
                ByVal Value As TaskContext)
Const ProcName As String = "ITask_TaskContext"
Set mTaskContext = Value
End Property

Private Property Get ITask_taskName() As String
ITask_taskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureBuilder Event Handlers
'@================================================================================

Private Sub mFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mFutureBuilder_Cancelled"
On Error GoTo Err

doCancel

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFetchInstrumentsTC Event Handlers
'@================================================================================

Private Sub mFetchInstrumentsTC_Completed( _
                ev As TaskCompletionEventData)
Const ProcName As String = "mFetchInstrumentsTC_Completed"
On Error GoTo Err

If ev.Cancelled Then
    doCancel
ElseIf ev.ErrorNumber <> 0 Then
    mFutureBuilder.Fail ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
    If Not mListener Is Nothing Then mListener.FetchFailed mFutureBuilder.cookie, ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
    mTaskContext.Finish Empty, True
Else
    Dim instruments As DataObjects
    Set instruments = ev.Result
    Set mInstrEnumerator = instruments.Enumerator
    
    mTaskContext.Continue
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not mListener Is Nothing Then mListener.FetchCompleted mFutureBuilder.cookie
mTaskContext.Finish Empty, False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ContractsFuture() As IFuture
Set ContractsFuture = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal Specifier As IContractSpecifier, _
                ByVal pListener As IContractFetchListener, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
Set mDB = db
Set mListener = pListener
Set mContractSpec = Specifier
mFutureBuilder.cookie = pCookie
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

If Not mFetchInstrumentsTC Is Nothing Then mFetchInstrumentsTC.CancelTask
mFutureBuilder.Cancel
If Not mListener Is Nothing Then mListener.FetchCancelled mFutureBuilder.cookie

mTaskContext.Finish Empty, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "logMessage"
gLogger.Log pMsg, pProcName, ModuleName, pLogLevel, pMsgQualifier
End Sub





