VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfilePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements IGenericTickListener
Implements ITickfileReplayMonitor
Implements Task

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "TickfilePlayer"

'@================================================================================
' Enums
'@================================================================================

Private Enum States
    NeedingTick
    FetchingTick
    HandlingTick
    PendingFireTick
    Finished
End Enum

Private Enum TickfileStates
    TickfileStateNotPlaying = 0
    TickfileStatePlaying = 1
    TickfileStatePaused = 2
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileManager                As TickFileManager
Private mTickfileSpec                   As ITickfileSpecifier
Private mTickScheduler                  As TickScheduler

Private mState                          As States

Private mPrimaryContractStore           As IContractStore
Private mSecondaryContractStore         As IContractStore

Private mCurrentTick                    As GenericTick

Private mReader                         As ITickfileReader

Private WithEvents mTickStreamBuilder   As TickStreamBuilder
Attribute mTickStreamBuilder.VB_VarHelpID = -1
Private mTickStream                     As TickStream

Private mTickfileState                  As TickfileStates

Private mTimestampAdjustmentStart       As Double
Private mTimestampAdjustmentEnd         As Double

Private mClock                          As Clock

Private mTaskContext                    As TaskContext

Private mEventsPlayed                   As Long
Private mPercentComplete                As Single

Private mReplaySpeed                    As Long

Private WithEvents mFutureWaiter        As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
Private mContract                       As IContract

Private mFirstTickTimeUTC               As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mTickfileState = TickfileStates.TickfileStateNotPlaying
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

mEventsPlayed = mEventsPlayed + 1

mCurrentTick = ev.Tick
mCurrentTick.Timestamp = adjustTimestamp(mCurrentTick.Timestamp)

mState = HandlingTick
If mTaskContext.State = TaskStateSuspended Then mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITickfileReplayMonitor Interface Members
'@================================================================================

Private Sub ITickfileReplayMonitor_NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "ITickfileReplayMonitor_NotifyEvent"
On Error GoTo Err

Set ev.Source = Me
mTickfileManager.NotifyEvent ev

Select Case ev.EventCode
Case TickfileEventCodes.TickfileEventFileDoesNotExist, _
        TickfileEventCodes.TickfileEventFileIsEmpty, _
        TickfileEventCodes.TickfileEventFileIsInvalid, _
        TickfileEventCodes.TickfileEventFileFormatNotSupported, _
        TickfileEventCodes.TickfileEventAmbiguousContractDetails, _
        TickfileEventCodes.TickfileEventNoContractDetails
    mTickfileManager.NotifySkipTickfile Me, gGetTickfileEventData(mTickfileManager, mTickfileSpec, Me), 0
Case TickfileEventCodes.TickfileEventDataSourceNotAvailable
    mTickfileManager.StopReplay
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileReplayMonitor_NotifyReady( _
                ByVal pSource As Object, _
                ByVal pTickfileSpec As ITickfileSpecifier, _
                Optional ByVal pContractSpec As IContractSpecifier)
Const ProcName As String = "ITickfileReplayMonitor_NotifyReady"
On Error GoTo Err

Set mTickfileSpec = pTickfileSpec
If Not mTickfileSpec.ContractFuture Is Nothing Then
    mFutureWaiter.Add mTickfileSpec.ContractFuture
ElseIf Not pContractSpec Is Nothing Then
    
    'now query the Contract Store(s) to see if we can
    ' determine the canonical contract for this contract spec
    mTickfileSpec.ContractFuture = getContractFuture(pContractSpec)

Else
    ' we have no way to determine what the contract is
    fireNotifyEvent TickfileEventCodes.TickfileEventNoContractDetails, "No contract details found for tickfile"
    mTickfileManager.NotifySkipTickfile Me, gGetTickfileEventData(mTickfileManager, mTickfileSpec, Me), 0
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileReplayMonitor_NotifyReplayComplete(ByVal pSource As Object, ByVal pEventsPlayed As Long)
Const ProcName As String = "ITickfileReplayMonitor_NotifyReplayComplete"
On Error GoTo Err

mState = Finished
mTaskContext.Continue
mReader.CloseInputFile

Dim lTickStreamBuilder As TickStreamBuilder
Set lTickStreamBuilder = mTickStreamBuilder
' ignore state change event
Set mTickStreamBuilder = Nothing

lTickStreamBuilder.Finish
mTickfileManager.NotifyTickfileCompleted Me, gGetTickfileEventData(mTickfileManager, mTickfileSpec, Me), mEventsPlayed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
End Sub

Private Sub Task_Run()
Const ProcName As String = "Task_Run"
On Error GoTo Err

'debug.print FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  Task_Run"

If mTaskContext.CancelPending Then
    StopReplay
    Exit Sub
End If

If mTickfileState = TickfileStates.TickfileStateNotPlaying Or _
    mTickfileState = TickfileStates.TickfileStatePaused _
Then
    mTaskContext.Suspend -1
    Exit Sub
End If

If mState = States.HandlingTick Then
    mState = States.PendingFireTick
    mTickScheduler.NotifyTick mCurrentTick, Me
    If Not mState = States.NeedingTick Then mTaskContext.Suspend -1
ElseIf mState = States.NeedingTick Then
    If Not mReader.FireNextTick Then
        mTaskContext.Suspend -1
        mState = States.FetchingTick
    End If
ElseIf mState = FetchingTick Then
    mTaskContext.Suspend -1
ElseIf mState = PendingFireTick Then
    mTaskContext.Suspend -1
ElseIf mState = Finished Then
    mTickfileState = TickfileStates.TickfileStateNotPlaying
    mTickScheduler.NotifyNoMoreTicks Me
    mTaskContext.Finish Empty, False
Else
    Assert True, "Invalid state"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    Err.Raise ev.Future.ErrorNumber, ev.Future.ErrorSource, ev.Future.ErrorMessage
ElseIf ev.Future.IsCancelled Then
    mTickfileManager.NotifySkipTickfile Me, gGetTickfileEventData(mTickfileManager, mTickfileSpec, Me), 0
Else
    Set mContract = ev.Future.Value
    checkReady
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTickStreamBuilder Event Handlers
'@================================================================================

Private Sub mTickStreamBuilder_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mTickStreamBuilder_StateChange"
On Error GoTo Err

If ev.State <> TickStreamStates.TickStreamStateFinished Then Exit Sub

mState = Finished
mTaskContext.Continue
mReader.CloseInputFile
mTickfileManager.NotifyTickfileCompleted Me, gGetTickfileEventData(mTickfileManager, mTickfileSpec, Me), mEventsPlayed

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let FirstTickTimeUTC(ByVal Value As Date)
Const ProcName As String = "FirstTickTimeUTC"
On Error GoTo Err

mFirstTickTimeUTC = Value
checkReady

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ReplayProgressTicks() As Long
ReplayProgressTicks = mReader.ReplayProgressTicks
End Property

Friend Property Get ReplayProgressPercent() As Long
ReplayProgressPercent = mReader.ReplayProgressPercent
End Property

Friend Property Let ReplaySpeed(ByVal Value As Long)
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

mReplaySpeed = Value
If Not mClock Is Nothing Then mClock.Rate = mReplaySpeed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ReplaySpeed() As Long
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

ReplaySpeed = mReplaySpeed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get State() As TickfileStates
Const ProcName As String = "State"
On Error GoTo Err

State = mTickfileState

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get TickStream() As TickStream
Set TickStream = mTickStream
End Property

Friend Property Get TickfileSizeBytes() As Long
Const ProcName As String = "TickfileSizeBytes"
On Error GoTo Err

If mReader Is Nothing Then
    TickfileSizeBytes = 0
Else
    TickfileSizeBytes = mReader.TickfileSizeBytes
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get TimeZone() As TimeZone
Set TimeZone = mClock.TimeZone
End Property

Friend Property Get Clock() As Clock
Set Clock = mClock
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pTickfileManager As TickFileManager, _
                ByVal pTickScheduler As TickScheduler, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pStreamId As Long, _
                ByVal pReplaySpeed As Long, _
                ByVal pTimestampAdjustmentStart As Double, _
                ByVal pTimestampAdjustmentEnd As Double)
Set mTickfileManager = pTickfileManager
Set mTickScheduler = pTickScheduler
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore
Set mTickStreamBuilder = CreateTickStreamBuilder(pStreamId)
Set mTickStream = mTickStreamBuilder.TickStream
ReplaySpeed = pReplaySpeed
mTimestampAdjustmentStart = pTimestampAdjustmentStart / 86400
mTimestampAdjustmentEnd = pTimestampAdjustmentEnd / 86400
End Sub

Friend Sub PauseReplay()
Const ProcName As String = "PauseReplay"
On Error GoTo Err

Assert mTickfileState = TickfileStatePlaying, "Pause when not playing"

mTickfileState = TickfileStates.TickfileStatePaused
mClock.PauseTime
mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function PrepareTickfile( _
                ByVal pTickfileReader As ITickfileReader, _
                ByVal pTickfileSpec As TickfileSpecifier) As Boolean
Const ProcName As String = "PrepareTickfile"
On Error GoTo Err

Set mReader = pTickfileReader
Set mTickfileSpec = pTickfileSpec
PrepareTickfile = mReader.PrepareTickfile(mTickfileSpec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ProcessTick(ByRef pTick As GenericTick)
Const ProcName As String = "ProcessTick"
On Error GoTo Err

mClock.SetTime pTick.Timestamp

mTickStreamBuilder.NotifyTick pTick
mState = States.NeedingTick
If mTaskContext.State = TaskStateSuspended Then
    'debug.print FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  Task.Continue"
    mTaskContext.Continue
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ResumeReplay()
Const ProcName As String = "ResumeReplay"
On Error GoTo Err
 
Assert mTickfileState = TickfileStatePaused, "Resume without Pause"

mTickfileState = TickfileStates.TickfileStatePlaying
mClock.ResumeTime
mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StartReplay()
Const ProcName As String = "StartReplay"
On Error GoTo Err

If mTickfileState = TickfileStates.TickfileStatePlaying Then Exit Sub

mTickfileState = TickfileStates.TickfileStatePlaying
mState = States.NeedingTick
StartTask Me, PriorityLow

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StopReplay()
Const ProcName As String = "StopReplay"
On Error GoTo Err

mState = Finished
If Not mTaskContext Is Nothing Then mTaskContext.Continue

If Not mReader Is Nothing Then mReader.CloseInputFile
Set mReader = Nothing

Dim lTickStreamBuilder As TickStreamBuilder
Set lTickStreamBuilder = mTickStreamBuilder
' ignore state change event
Set mTickStreamBuilder = Nothing

lTickStreamBuilder.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function adjustTimestamp(ByVal Timestamp As Date) As Date
Const ProcName As String = "adjustTimestamp"
On Error GoTo Err

adjustTimestamp = Timestamp + _
                mTimestampAdjustmentStart + _
                (mTimestampAdjustmentEnd - mTimestampAdjustmentStart) * (mPercentComplete / 100)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub fireNotifyEvent(ByVal pEventCode As Long, ByVal pEventMessage As String)
Dim lEv As NotificationEventData

Set lEv.Source = Me
lEv.EventCode = pEventCode
lEv.EventMessage = pEventMessage & ": " & _
                gTickfileSpecifierToString(mTickfileSpec)
mTickfileManager.NotifyEvent lEv
End Sub

Private Function getContractFuture(ByVal pContractSpec As IContractSpecifier) As IFuture
Const ProcName As String = "getContractFuture"
On Error GoTo Err

If mPrimaryContractStore Is Nothing Then
    fireNotifyEvent TickfileEventCodes.TickfileEventNoContractDetails, "No contract store is available"
    mTickfileManager.NotifySkipTickfile Me, gGetTickfileEventData(mTickfileManager, mTickfileSpec, Me), 0
    Exit Function
End If
    
Set getContractFuture = FetchContract(pContractSpec, mPrimaryContractStore, mSecondaryContractStore)
mFutureWaiter.Add getContractFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub checkReady()
Const ProcName As String = "checkReady"
On Error GoTo Err

If mContract Is Nothing Then Exit Sub

If mClock Is Nothing Then Set mClock = CreateSimulatedClock(mReplaySpeed, mContract.TimezoneName)

If mFirstTickTimeUTC = 0 Then
    mTickfileManager.NotifyReady Me
    Exit Sub
End If

mClock.SetTimeUTC mFirstTickTimeUTC
mTickfileManager.NotifyTickStreamStarted gGetTickfileEventData(mTickfileManager, mTickfileSpec, Me), mClock

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupClock()
Const ProcName As String = "setupClock"
On Error GoTo Err

Set mClock = CreateSimulatedClock(mReplaySpeed, mContract.TimezoneName)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


