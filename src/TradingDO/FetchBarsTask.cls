VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FetchBarsTask"
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


Private Const ModuleName                    As String = "FetchBarsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mDB                                 As TradingDB

Private mSpecifier                          As IContractSpecifier
Private mBarTimePeriod                      As TimePeriod
Private mMaxNumberOfBars                    As Long
Private mFromDate                           As Date
Private mToDate                             As Date
Private mCustomSessionStartTime             As Date
Private mCustomSessionEndTime               As Date
Private mIncludeBarsOutsideSession          As Boolean
Private mBarType                            As BarTypes

Private mTaskContext                        As TaskContext

Private WithEvents mInstrumentTaskController    As TaskController
Attribute mInstrumentTaskController.VB_VarHelpID = -1
                
Private mInstrument                         As instrument

Private mListener                           As IBarFetchListener
Private mCookie                             As Variant

Private mBarsBuilder                        As BarsBuilder
Private mSessionBuilder                     As SessionBuilder

Private WithEvents mFutureWaiter            As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
Private mFutureBuilder                      As New FutureBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

Private Sub Class_Terminate()
Debug.Print "FetchBarsTask terminated"
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
Const ProcName As String = "ITask_Cancel"
On Error GoTo Err

mFutureBuilder.Cancel
If Not mInstrumentTaskController Is Nothing Then mInstrumentTaskController.CancelTask
mTaskContext.Finish Empty, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    mFutureBuilder.Cancel
    mTaskContext.Finish Empty, True
    Exit Sub
End If

If mInstrument Is Nothing Then
    gLogger.Log "FetchBarsTask started for " & mSpecifier.ToString, ProcName, ModuleName
    Set mInstrumentTaskController = mDB.InstrumentFactory.LoadBySpecifierAsync( _
                                            mSpecifier, _
                                            , _
                                            "LoadBySpecifier/" & mSpecifier.ToString & "(" & gGetSequenceNumber & ")")
    mTaskContext.Suspend -1
Else
    mFutureWaiter.Add mInstrument.FetchBarsAsync(mBarTimePeriod, _
                                                        mMaxNumberOfBars, _
                                                        mFromDate, _
                                                        mToDate, _
                                                        mCustomSessionStartTime, _
                                                        mCustomSessionEndTime, _
                                                        mIncludeBarsOutsideSession, _
                                                        mBarType, _
                                                        mListener, _
                                                        mCookie)
    mTaskContext.Suspend -1
End If

Exit Sub

Err:

If Err.Number = 3265 Then   ' "Item cannot be found in the collection corresponding to the requested name or ordinal"
    ' this sometimes happens with the first async operation for a long time - retry it
    ' up to 5 times
    Static sRetryCount As Long
    
    If sRetryCount < 5 Then
        Set mInstrument = Nothing
        sRetryCount = sRetryCount + 1
        mTaskContext.Suspend 10
        Exit Sub
    End If
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal value As TaskContext)
Set mTaskContext = value
End Property

Private Property Get ITask_taskName() As String
ITask_taskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsPending Then

ElseIf ev.Future.IsCancelled Then
    mFutureBuilder.Cancel
    mTaskContext.Finish Empty, True
ElseIf ev.Future.IsFaulted Then
    gLogger.Log "FetchBarsTask failed for " & mSpecifier.ToString & vbCrLf & _
                                "Error " & ev.Future.ErrorNumber & ": " & ev.Future.ErrorMessage, ProcName, ModuleName, LogLevelSevere
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
    mTaskContext.Error ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
Else
    mFutureBuilder.value = ev.Future.value
    mFutureBuilder.Complete
    gLogger.Log "FetchBarsTask retrieved bars for " & mSpecifier.ToString, ProcName, ModuleName
    mTaskContext.Finish Empty, False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mInstrumentTaskController Event Handlers
'@================================================================================

Private Sub mInstrumentTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mInstrumentTaskController_Completed"
On Error GoTo Err

If ev.Cancelled Then
    mFutureBuilder.Cancel
    mTaskContext.Finish Empty, True
ElseIf ev.ErrorNumber <> 0 Then
    mFutureBuilder.Fail ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
    mTaskContext.Error ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
ElseIf IsEmpty(ev.Result) Then
    gLogger.Log "FetchBarsTask cannot retrieve instrument object for " & mSpecifier.ToString, ProcName, ModuleName
    mFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, "Specified object does not exist", ""
    mTaskContext.Error ErrorCodes.ErrIllegalArgumentException, "Specified object does not exist"
Else
    Set mInstrument = ev.Result
    gLogger.Log "FetchBarsTask retrieved instrument object for " & mSpecifier.ToString, ProcName, ModuleName
    mTaskContext.Continue
End If
Set mInstrumentTaskController = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get BarsFuture() As IFuture
Const ProcName As String = "BarsFuture"
On Error GoTo Err

Set BarsFuture = mFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal Specifier As IContractSpecifier, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal maxNumberOfBars As Long, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                ByVal customSessionStartTime As Date, _
                ByVal customSessionEndTime As Date, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal barType As BarTypes, _
                ByVal pListener As IBarFetchListener, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mDB = db
Set mSpecifier = Specifier
Set mBarTimePeriod = BarTimePeriod
mMaxNumberOfBars = maxNumberOfBars
mFromDate = fromDate
mToDate = toDate
mCustomSessionStartTime = customSessionStartTime
mCustomSessionEndTime = customSessionEndTime
mIncludeBarsOutsideSession = includeBarsOutsideSession
mBarType = barType
Set mListener = pListener
gSetVariant mCookie, pCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


