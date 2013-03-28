VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MergedClrTask"
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
    Paused
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MergedClrTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mTickfileManager                            As TickFileManager
Private mTickfileSpecifiers                         As TickFileSpecifiers

Private mTickscheduler                              As TickScheduler

Private mEn                                         As Enumerator

Private mPlayer                                     As TickfilePlayer

Private mState                                      As States

Private mReplaySpeed                                As Long

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

Private Sub ITickfileControllerTask_Initialise(ByVal pTickfileManager As TickFileManager, ByVal pTickfileSpecifiers As TickFileSpecifiers, ByVal pReplaySpeed As Long)
Initialise pTickfileManager, pTickfileSpecifiers, pReplaySpeed
End Sub

Private Sub ITickfileControllerTask_NotifyReady()
NotifyReady
End Sub

Private Sub ITickfileControllerTask_NotifyTickfileCompleted()
NotifyTickfileCompleted
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

Private Sub ITickfileControllerTask_Skip()
Skip
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

cancelReplay
mTaskContext.Finish Empty, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Task_Run()
Const ProcName As String = "Task_Run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    cancelReplay
    mTaskContext.Finish Empty, True
    Exit Sub
End If

If mState = Inactive Then
    If Not mEn.MoveNext Then
        mTaskContext.Finish Empty, False
        Exit Sub
    End If
    
    mState = AwaitingReady
    Set mTickscheduler = New TickScheduler
    mTickscheduler.Initialise mReplaySpeed

    Set mPlayer = mTickfileManager.SetupTickfilePlayer(mEn.Current, mTickscheduler)
    If mPlayer Is Nothing Then
        mState = Inactive
    ElseIf mState <> Ready Then
        mTaskContext.Suspend -1
    End If
ElseIf mState = Ready Then
    Select Case mTickfileManager.QueryPlay(mEn.Current, mPlayer)
    Case ReplayContinueModes.ReplayCancelAllFiles
        mPlayer.StopReplay
        mState = Inactive
        mTaskContext.Finish Empty, True
    Case ReplayContinueModes.ReplayCancelThisFile
        mTickfileManager.NotifySkipTickfile gGetTickfileEventData(mTickfileManager, mEn.Current, mPlayer), 0
        mState = Inactive
    Case ReplayContinueModes.ReplayContinue
        mState = Playing
        mTickscheduler.AddPlayer mPlayer
        mTickscheduler.StartScheduling
        mPlayer.StartReplay
        mTaskContext.Suspend -1
    End Select
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
If Not mTickscheduler Is Nothing Then mTickscheduler.ReplaySpeed = mReplaySpeed

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
                ByVal pReplaySpeed As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTickfileManager = pTickfileManager
Set mTickfileSpecifiers = pTickfileSpecifiers
Set mEn = mTickfileSpecifiers.Enumerator
mReplaySpeed = pReplaySpeed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyReady()
Const ProcName As String = "NotifyReady"
On Error GoTo Err

mState = Ready
If mTaskContext.State = TaskStateSuspended Then mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTickfileCompleted()
Const ProcName As String = "NotifyTickfileCompleted"
On Error GoTo Err

Set mPlayer = Nothing
mState = Inactive
mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PauseReplay()
Const ProcName As String = "PauseReplay"
On Error GoTo Err

mState = Paused
If Not mPlayer Is Nothing Then mPlayer.PauseReplay
If mTaskContext.State = TaskStateRunnable Then mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ResumeReplay()
Const ProcName As String = "ResumeReplay"
On Error GoTo Err

If Not mPlayer Is Nothing Then
    mState = Playing
    mPlayer.ResumeReplay
Else
    mState = Inactive
End If
mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Skip()
Const ProcName As String = "Skip"
On Error GoTo Err

If Not mPlayer Is Nothing Then mPlayer.StopReplay
Set mPlayer = Nothing
mState = Inactive
If mTaskContext.State = TaskStateSuspended Then mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StopReplay()
Const ProcName As String = "StopReplay"
On Error GoTo Err

If Not mPlayer Is Nothing Then mPlayer.StopReplay
mTaskContext.Finish Empty, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub cancelReplay()
If Not mPlayer Is Nothing Then mPlayer.StopReplay
End Sub




