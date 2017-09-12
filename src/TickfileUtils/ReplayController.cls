VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReplayController"
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

Event NotifyEvent(ev As NotificationEventData)

Event ReplayPaused()

Event ReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Long)

Event ReplayResumed()

Event ReplayStarted()
Event ReplayStopped()

Event TickfileCompleted( _
                ByRef ev As TickfileEventData, _
                ByVal pEventsPlayed As Long)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ReplayController"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileSpecifiers                         As TickFileSpecifiers
Private mTickfileIndex                              As Long

Private mTickfileStore                              As ITickfileStore
Private mPrimaryContractStore                       As IContractStore
Private mSecondaryContractStore                     As IContractStore

Private mReplaySpeed                                As Long
Private mReplayProgressEventInterval                As Long
Private mTimestampAdjustmentStart                   As Double
Private mTimestampAdjustmentEnd                     As Double

Private mReplayInProgress                           As Boolean
Private mPaused                                     As Boolean

Private mPlayers()                                  As TickfilePlayer

Private mListeners                                  As New Listeners

Private mNumberOfPlayersAwaitingReady               As Long
Private mNumberOfPlayersAwaitingPrepare             As Long
Private mNumberOfPlayersNotCompleted                As Long

Private mTickScheduler                              As TickScheduler

Private mTaskContext                                As TaskContext

Private mUnableToReplay                             As Boolean

Private mIsReplaySequential                         As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
Const ProcName As String = "ITask_Cancel"
On Error GoTo Err

StopReplay

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Static sStarted As Boolean
If Not sStarted Then
    sStarted = True
    gLogger.Log "Tickfile replay started", ProcName, ModuleName
    RaiseEvent ReplayStarted
End If

If mNumberOfPlayersNotCompleted = 0 Then
    terminateReplay
ElseIf mNumberOfPlayersAwaitingReady = 0 And mNumberOfPlayersAwaitingPrepare = 0 Then
    startScheduler
Else
    waitForPlayers
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ReplayInProgress() As Boolean
ReplayInProgress = mReplayInProgress
End Property

Public Property Let ReplaySpeed(ByVal Value As Long)
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

mReplaySpeed = Value

Dim i As Long
For i = 0 To UBound(mPlayers)
    mPlayers(i).ReplaySpeed = mReplaySpeed
Next

If Not mTickScheduler Is Nothing Then mTickScheduler.ReplaySpeed = mReplaySpeed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ReplaySpeed() As Long
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

ReplaySpeed = mReplaySpeed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SequentialTickFileSpecifiers(Value As TickFileSpecifiers)
Const ProcName As String = "SequentialTickFileSpecifiers"
On Error GoTo Err

Assert mTickfileSpecifiers Is Nothing, "TickfileSpecifiers or SequentialTickFileSpecifiers has already been set"
AssertArgument Value.Count > 0, "Value has no members"

Set mTickfileSpecifiers = Value.Clone

mIsReplaySequential = True

ReDim mPlayers(0) As TickfilePlayer

Set mTickScheduler = New TickScheduler
mTickScheduler.Initialise Me, mReplaySpeed, mReplayProgressEventInterval, True

Set mPlayers(0) = setupTickfilePlayer(0, mTickScheduler)
mTickfileIndex = 1
prepareTickfilePlayer mPlayers(0), mTickfileSpecifiers(mTickfileIndex), 0

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property


Public Property Let TickFileSpecifiers(Value As TickFileSpecifiers)
Const ProcName As String = "TickfileSpecifiers"
On Error GoTo Err

Assert mTickfileSpecifiers Is Nothing, "TickfileSpecifiers or SequentialTickFileSpecifiers has already been set"
AssertArgument Value.Count > 0, "Value has no members"

Set mTickfileSpecifiers = Value.Clone

ReDim mPlayers(mTickfileSpecifiers.Count - 1) As TickfilePlayer

Set mTickScheduler = New TickScheduler
mTickScheduler.Initialise Me, mReplaySpeed, mReplayProgressEventInterval, IIf(mTickfileSpecifiers.Count = 1, True, False)

Dim i As Long
For i = 0 To UBound(mPlayers)
    Set mPlayers(i) = setupTickfilePlayer(i, mTickScheduler)
    prepareTickfilePlayer mPlayers(i), mTickfileSpecifiers(i + 1), i
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickFileSpecifiers() As TickFileSpecifiers
Const ProcName As String = "TickfileSpecifiers"
On Error GoTo Err

Set TickFileSpecifiers = mTickfileSpecifiers

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickStream(ByVal pIndex As Long) As ITickStream
Const ProcName As String = "TickStream"
On Error GoTo Err

Assert Not mTickfileSpecifiers Is Nothing, "TickfileSpecifiers has not yet been set"

If Not mPlayers(pIndex) Is Nothing Then Set TickStream = mPlayers(pIndex).TickStream

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimestampAdjustmentStart() As Double
Const ProcName As String = "TimestampAdjustmentStart"
On Error GoTo Err

TimestampAdjustmentStart = mTimestampAdjustmentStart

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimestampAdjustmentEnd() As Double
Const ProcName As String = "TimestampAdjustmentEnd"
On Error GoTo Err

TimestampAdjustmentEnd = mTimestampAdjustmentEnd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddTickfileReplayListener(ByVal pListener As ITickfileReplayListener)
Const ProcName As String = "AddTickfileReplayListener"
On Error GoTo Err

mListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Intialise( _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pReplaySpeed As Long, _
                ByVal pTimestampAdjustmentStart As Double, _
                ByVal pTimestampAdjustmentEnd As Double, _
                ByVal pReplayProgressEventInterval As Long)
Const ProcName As String = "Intialise"
On Error GoTo Err

Set mTickfileStore = pTickfileStore
Set mPrimaryContractStore = pPrimaryContractStore
Set pSecondaryContractStore = pSecondaryContractStore
mReplaySpeed = pReplaySpeed
mTimestampAdjustmentStart = pTimestampAdjustmentStart
mTimestampAdjustmentEnd = pTimestampAdjustmentEnd

mReplayProgressEventInterval = pReplayProgressEventInterval

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCantPlay(ByVal pPlayer As TickfilePlayer)
Const ProcName As String = "NotifyCantPlay"
On Error GoTo Err

mUnableToReplay = True
mNumberOfPlayersNotCompleted = mNumberOfPlayersNotCompleted - 1
If mReplayInProgress Then mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "NotifyEvent"
On Error GoTo Err

Set ev.Source = Me
fireNotifyEvent ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyProgress(ByVal pTimestamp As Date, ByVal pEventsPlayed As Long, ByVal pPercentComplete As Long)
Const ProcName As String = "NotifyProgress"
On Error GoTo Err

fireProgress pTimestamp, pEventsPlayed, pPercentComplete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyReady(ByVal pPlayer As TickfilePlayer, ByVal pTickfileSpec As TickfileSpecifier)
Const ProcName As String = "NotifyReady"
On Error GoTo Err

gLogger.Log "Tickfile player ready", ProcName, ModuleName, , pTickfileSpec.ToString

Assert mNumberOfPlayersAwaitingReady > 0, "Ready notified when none expected"
mNumberOfPlayersAwaitingReady = mNumberOfPlayersAwaitingReady - 1

If mReplayInProgress Then mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTickfileCompleted(ByVal pPlayer As TickfilePlayer, ByRef ev As TickfileEventData, ByVal pEventsPlayed As Long)
Const ProcName As String = "NotifyTickfileCompleted"
On Error GoTo Err

gLogger.Log "Tickfile player completed", ProcName, ModuleName, , ev.Specifier.ToString

fireTickfileCompleted ev, pEventsPlayed

If mIsReplaySequential Then
    mTickfileIndex = mTickfileIndex + 1
    If mTickfileIndex <= mTickfileSpecifiers.Count Then
        prepareTickfilePlayer mPlayers(0), mTickfileSpecifiers(mTickfileIndex), 0
    Else
        If pPlayer.PlayerState <> PlayerStateCreated Then mTickScheduler.RemovePlayer pPlayer
        mNumberOfPlayersNotCompleted = mNumberOfPlayersNotCompleted - 1
    End If
Else
    If pPlayer.PlayerState <> PlayerStateCreated Then mTickScheduler.RemovePlayer pPlayer
    pPlayer.Finish
    mNumberOfPlayersNotCompleted = mNumberOfPlayersNotCompleted - 1
End If

If mReplayInProgress Then mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTickfilePrepared( _
                ByVal pSource As Object, _
                ByVal pTickfileSpec As ITickfileSpecifier)
Const ProcName As String = "NotifyTickfilePrepared"
On Error GoTo Err

mNumberOfPlayersAwaitingReady = mNumberOfPlayersAwaitingReady + 1
Assert mNumberOfPlayersAwaitingPrepare > 0, "Prepared notified when none expected"
mNumberOfPlayersAwaitingPrepare = mNumberOfPlayersAwaitingPrepare - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PauseReplay()
Const ProcName As String = "PauseReplay"
On Error GoTo Err

Assert mReplayInProgress, "Replay has not yet been started"

mPaused = True

mTickScheduler.PauseReplay

Dim i As Long
For i = 0 To UBound(mPlayers)
    If mPlayers(i).TickfileState = TickfileStatePlaying Then mPlayers(i).PauseReplay
Next

RaiseEvent ReplayPaused

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveTickfileReplayListener(ByVal pListener As ITickfileReplayListener)
Const ProcName As String = "RemoveTickfileReplayListener"
On Error GoTo Err

mListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ResumeReplay()
Const ProcName As String = "ResumeReplay"
On Error GoTo Err

Assert mReplayInProgress, "Replay has not yet been started"
Assert mPaused, "Replay has not been paused"

mPaused = False

Dim i As Long
For i = 0 To UBound(mPlayers)
    If mPlayers(i).TickfileState = TickfileStatePaused Then mPlayers(i).ResumeReplay
Next

mTickScheduler.ResumeReplay

RaiseEvent ReplayResumed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function StartReplay() As TaskController
Const ProcName As String = "StartReplay"
On Error GoTo Err

Assert Not mReplayInProgress, "Replay has already been started"
    
mReplayInProgress = True
Set StartReplay = StartTask(Me, PriorityNormal)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub StopReplay()
Const ProcName As String = "StopReplay"
On Error GoTo Err

Assert mReplayInProgress, "Replay has not yet been started"
mReplayInProgress = False

Dim i As Long
For i = 0 To UBound(mPlayers)
    mPlayers(i).Finish
Next

mTickScheduler.StopScheduling
mTaskContext.Finish Empty, True

RaiseEvent ReplayStopped

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireNotifyEvent(ev As NotificationEventData)
Const ProcName As String = "fireNotifyEvent"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mListeners.Valid Then
    sInit = True
    sSomeListeners = mListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ITickfileReplayListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyEvent ev
    Next
End If

RaiseEvent NotifyEvent(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireProgress(ByVal pTimestamp As Date, ByVal pEventsPlayed As Long, ByVal pPercentComplete As Long)
Const ProcName As String = "fireProgress"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mListeners.Valid Then
    sInit = True
    sSomeListeners = mListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ITickfileReplayListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.ReplayProgress pTimestamp, pEventsPlayed, pPercentComplete
    Next
End If

RaiseEvent ReplayProgress(pTimestamp, pEventsPlayed, pPercentComplete)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireTickfileCompleted(ByRef ev As TickfileEventData, ByVal pEventsPlayed As Long)
Const ProcName As String = "fireTickfileCompleted"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mListeners.Valid Then
    sInit = True
    sSomeListeners = mListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ITickfileReplayListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.TickfileCompleted ev, pEventsPlayed
    Next
End If

RaiseEvent TickfileCompleted(ev, pEventsPlayed)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub prepareTickfilePlayer(ByVal pPlayer As TickfilePlayer, ByVal tickfileSpec As TickfileSpecifier, ByVal pStreamId As Long)
Const ProcName As String = "prepareTickfilePlayer"
On Error GoTo Err

Dim lTickfileReader As ITickfileReader
Set lTickfileReader = mTickfileStore.CreateTickfileReader(pPlayer, pPlayer, pStreamId)

Assert Not lTickfileReader Is Nothing, "Failed to create TickfileReader"

If pPlayer.PrepareTickfile(lTickfileReader, tickfileSpec) Then
    gLogger.Log "Tickfile player awaiting ready", ProcName, ModuleName, , tickfileSpec.ToString
    mNumberOfPlayersAwaitingReady = mNumberOfPlayersAwaitingReady + 1
Else
    gLogger.Log "Tickfile player awaiting prepare", ProcName, ModuleName, , tickfileSpec.ToString
    mNumberOfPlayersAwaitingPrepare = mNumberOfPlayersAwaitingPrepare + 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupTickfilePlayer( _
                ByVal pStreamId As Long, _
                ByVal pTickScheduler As TickScheduler) As TickfilePlayer
Const ProcName As String = "setupTickfilePlayer"
On Error GoTo Err

Dim lTickfilePlayer As TickfilePlayer
Set lTickfilePlayer = New TickfilePlayer
lTickfilePlayer.Initialise Me, _
                    pTickScheduler, _
                    mPrimaryContractStore, _
                    mSecondaryContractStore, _
                    pStreamId, _
                    mReplaySpeed, _
                    mTimestampAdjustmentStart, _
                    mTimestampAdjustmentEnd

mNumberOfPlayersNotCompleted = mNumberOfPlayersNotCompleted + 1

Set setupTickfilePlayer = lTickfilePlayer

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub startScheduler()
Const ProcName As String = "startScheduler"
On Error GoTo Err

gLogger.Log "Tickfile replay scheduling started", ProcName, ModuleName
mTickScheduler.StartScheduling
mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub terminateReplay()
Const ProcName As String = "terminateReplay"
On Error GoTo Err

gLogger.Log "Tickfile replay scheduling stopped", ProcName, ModuleName
mTickScheduler.StopScheduling
If mUnableToReplay Then
    gLogger.Log "Tickfile replay completed with errors: one of more files was invalid", ProcName, ModuleName
    mTaskContext.Error ErrorCodes.ErrIllegalArgumentException, "One or more files was invalid"
Else
    gLogger.Log "Tickfile replay scheduling completed successfully", ProcName, ModuleName
    mTaskContext.Finish Empty, False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub waitForPlayers()
Const ProcName As String = "waitForPlayers"
On Error GoTo Err

gLogger.Log "Waiting for tickfile players: " & _
            mNumberOfPlayersAwaitingPrepare & _
            " to prepare and " & _
            mNumberOfPlayersAwaitingReady & _
            " to be ready", _
            ProcName, ModuleName
mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

