VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SequentialControllerTask"
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

Private Const ModuleName                            As String = "SequentialControllerTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mTickfileManager                            As TickFileManager
Private mTickfileSpecifiers                         As TickFileSpecifiers

Private mTickfileStore                              As ITickfileStore

Private mDataConsumer                               As IGenericTickListener

Private mPrimaryContractStore                       As IContractStore
Private mSecondaryContractStore                     As IContractStore

Private mEn                                         As Enumerator

Private mPlayer                                     As TickfilePlayer

Private mState                                      As States

'@================================================================================
' Class Event Handlers
'@================================================================================

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

If Not mEn.MoveNext Then
    mTaskContext.Finish Empty, False
    Exit Sub
End If

If mState = Inactive Then
    mState = AwaitingReady
    Set mPlayer = mTickfileManager.SetupTickfilePlayer(mEn.Current)
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
        mState = Inactive
    Case ReplayContinueModes.ReplayContinue
        mState = Playing
        mPlayer.StartReplay
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

'@================================================================================
' Methods
'@================================================================================
Friend Sub Initialise( _
                ByVal pTickfileManager As TickFileManager, _
                ByVal pTickfileSpecifiers As TickFileSpecifiers, _
                ByVal pDataConsumer As IGenericTickListener, _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTickfileManager = pTickfileManager
Set mTickfileSpecifiers = pTickfileSpecifiers
Set mEn = mTickfileSpecifiers.Enumerator

Set mDataConsumer = pDataConsumer
Set mTickfileStore = pTickfileStore
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore

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

mPlayer.StopReplay
Set mPlayer = Nothing
mState = Inactive
If mTaskContext.State = TaskStateSuspended Then mTaskContext.Continue

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


