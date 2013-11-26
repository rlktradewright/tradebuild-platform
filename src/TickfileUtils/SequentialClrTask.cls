VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SequentialClrTask"
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

Implements ITickfileControllerTask
Implements Task

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum States
    Inactive
    AwaitingReady
    Ready
    Playing
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "SequentialClrTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mTickfileManager                            As TickFileManager
Private mTickfileSpecifiers                         As TickFileSpecifiers

Private mTickScheduler                              As TickScheduler

Private mEn                                         As Enumerator

Private mPlayer                                     As TickfilePlayer

Private mState                                      As States
Private mPaused                                     As Boolean

Private mReplaySpeed                                As Long
Private mReplayProgressEventInterval                As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITickfileControllerTask Interface Members
'@================================================================================

Private Property Get ITickfileControllerTask_CurrentPlayer() As TickfilePlayer
Set ITickfileControllerTask_CurrentPlayer = CurrentPlayer
End Property

Private Property Get ITickfileControllerTask_CurrentTickfileSpec() As ITickfileSpecifier
Set ITickfileControllerTask_CurrentTickfileSpec = CurrentTickfileSpec
End Property

Private Sub ITickfileControllerTask_Initialise( _
                ByVal pTickfileManager As TickFileManager, _
                ByVal pTickfileSpecifiers As TickFileSpecifiers, _
                ByVal pReplaySpeed As Long, _
                ByVal pReplayProgressEventInterval As Long)
Initialise pTickfileManager, pTickfileSpecifiers, pReplaySpeed, pReplayProgressEventInterval
End Sub

Private Sub ITickfileControllerTask_NotifyReady(ByVal pPlayer As TickfilePlayer)
NotifyReady pPlayer
End Sub

Private Sub ITickfileControllerTask_NotifySkip(ByVal pPlayer As TickfilePlayer)
NotifySkip pPlayer
End Sub

Private Sub ITickfileControllerTask_NotifyTickfileCompleted(ByVal pPlayer As TickfilePlayer)
NotifyTickfileCompleted pPlayer
End Sub

Private Sub ITickfileControllerTask_PauseReplay()
PauseReplay
End Sub

Private Property Let ITickfileControllerTask_ReplaySpeed(ByVal Value As Long)
ReplaySpeed = Value
End Property

Private Sub ITickfileControllerTask_ResumeReplay()
ResumeReplay
End Sub

Private Sub ITickfileControllerTask_StopReplay()
StopReplay
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
Const ProcName As String = "Task_Cancel"
On Error GoTo Err

StopReplay

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Task_Run()
Const ProcName As String = "Task_Run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    StopReplay
    Exit Sub
End If

If mPaused Then Exit Sub

If mState = Inactive Then
    If Not mEn.MoveNext Then
        mTaskContext.Finish Empty, False
        Exit Sub
    End If
    
    mState = AwaitingReady
    Set mTickScheduler = New TickScheduler
    mTickScheduler.Initialise mTickfileManager, mReplaySpeed, mReplayProgressEventInterval, True

    Set mPlayer = mTickfileManager.SetupTickfilePlayer(mEn.Current, mTickScheduler)
    If mPlayer Is Nothing Then
        mState = Inactive
    ElseIf mState = AwaitingReady Then
        mTaskContext.Suspend -1
    End If
ElseIf mState = Ready Then
    Select Case mTickfileManager.QueryPlay(mEn.Current, mPlayer)
    Case ReplayContinueModes.ReplayCancelAllFiles
        mPlayer.StopReplay
        mState = Inactive
        mTaskContext.Finish Empty, True
    Case ReplayContinueModes.ReplayCancelThisFile
        mTickfileManager.NotifySkipTickfile mPlayer, gGetTickfileEventData(mTickfileManager, mEn.Current, mPlayer), 0
        mState = Inactive
    Case ReplayContinueModes.ReplayContinue
        mState = Playing
        mTickScheduler.AddPlayer mPlayer
        mTickScheduler.StartScheduling
        mPlayer.StartReplay
        mTaskContext.Suspend -1
    End Select
ElseIf mState = AwaitingReady Then
    mTaskContext.Suspend -1
ElseIf mState = Playing Then
    mTaskContext.Suspend -1
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
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get CurrentTickfileSpec() As ITickfileSpecifier
Set CurrentTickfileSpec = mEn.Current
End Property

Friend Property Get CurrentPlayer() As TickfilePlayer
Set CurrentPlayer = mPlayer
End Property

Friend Property Let ReplaySpeed(ByVal Value As Long)
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

mReplaySpeed = Value
If Not mPlayer Is Nothing Then mPlayer.ReplaySpeed = mReplaySpeed
If Not mTickScheduler Is Nothing Then mTickScheduler.ReplaySpeed = mReplaySpeed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================
Friend Sub Initialise( _
                ByVal pTickfileManager As TickFileManager, _
                ByVal pTickfileSpecifiers As TickFileSpecifiers, _
                ByVal pReplaySpeed As Long, _
                ByVal pReplayProgressEventInterval As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTickfileManager = pTickfileManager
Set mTickfileSpecifiers = pTickfileSpecifiers
Set mEn = mTickfileSpecifiers.Enumerator
mReplaySpeed = pReplaySpeed
mReplayProgressEventInterval = pReplayProgressEventInterval

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyReady(ByVal pPlayer As TickfilePlayer)
Const ProcName As String = "NotifyReady"
On Error GoTo Err

mState = Ready
If Not mPaused And mTaskContext.State = TaskStateSuspended Then mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifySkip(ByVal pPlayer As TickfilePlayer)
Const ProcName As String = "Skip"
On Error GoTo Err

If Not mPlayer Is Nothing Then mPlayer.StopReplay
Set mPlayer = Nothing

If Not mTickScheduler Is Nothing Then mTickScheduler.StopScheduling
Set mTickScheduler = Nothing

mState = Inactive
If Not mPaused And mTaskContext.State = TaskStateSuspended Then mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTickfileCompleted(ByVal pPlayer As TickfilePlayer)
Const ProcName As String = "NotifyTickfileCompleted"
On Error GoTo Err

Set mPlayer = Nothing
If Not mTickScheduler Is Nothing Then mTickScheduler.StopScheduling
Set mTickScheduler = Nothing

mState = Inactive
mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PauseReplay()
Const ProcName As String = "PauseReplay"
On Error GoTo Err

mPaused = True
If Not mTickScheduler Is Nothing Then mTickScheduler.PauseReplay
If Not mPlayer Is Nothing Then mPlayer.PauseReplay
If mTaskContext.State = TaskStateRunnable Then mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ResumeReplay()
Const ProcName As String = "ResumeReplay"
On Error GoTo Err

mPaused = False
If Not mPlayer Is Nothing Then mPlayer.ResumeReplay
If Not mTickScheduler Is Nothing Then mTickScheduler.ResumeReplay
mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StopReplay()
Const ProcName As String = "StopReplay"
On Error GoTo Err

If Not mPlayer Is Nothing Then mPlayer.StopReplay
Set mPlayer = Nothing

If Not mTickScheduler Is Nothing Then mTickScheduler.StopScheduling
Set mTickScheduler = Nothing

mTaskContext.Finish Empty, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



