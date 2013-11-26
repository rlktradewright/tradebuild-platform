VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickFileManager"
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

'@================================================================================
' Events
'@================================================================================

Event NotifyEvent(ev As NotificationEventData)

Event QueryReplayNextTickfile( _
                ByRef ev As TickfileEventData, _
                ByRef pContinueMode As ReplayContinueModes)

Event ReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Long)

Event TickStreamStarted( _
                ByRef ev As TickfileEventData, _
                ByVal pClock As Clock)

Event TickfileCompleted( _
                ByRef ev As TickfileEventData, _
                ByVal pEventsPlayed As Long)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "TickFileManager"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileStore                  As ITickfileStore

Private mTickfileSpecifiers             As TickFileSpecifiers

Private mReplayInProgress               As Boolean

Private mReplaySpeed                    As Long
Private mReplayProgressEventInterval    As Long
Private mTimestampAdjustmentStart       As Double
Private mTimestampAdjustmentEnd         As Double

Private mPrimaryContractStore           As IContractStore
Private mSecondaryContractStore         As IContractStore

Private mControllerTask                 As ITickfileControllerTask

Private mMerge                          As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get NumberOfInputTickfiles() As Long
Const ProcName As String = "NumberOfInputTickfiles"
On Error GoTo Err

NumberOfInputTickfiles = mTickfileSpecifiers.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ReplayInProgress() As Boolean
ReplayInProgress = mReplayInProgress
End Property

Public Property Get ReplayProgressEventInterval() As Long
Const ProcName As String = "ReplayProgressEventInterval"
On Error GoTo Err

ReplayProgressEventInterval = mReplayProgressEventInterval

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ReplaySpeed(ByVal Value As Long)
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

mReplaySpeed = Value
If Not mControllerTask Is Nothing Then mControllerTask.ReplaySpeed = mReplaySpeed

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

Public Property Let TickfileSpecifier(ByVal Value As TickfileSpecifier)
Const ProcName As String = "TickfileSpecifier"
On Error GoTo Err

Assert Not mReplayInProgress, "Tickfile is already replaying"

Set mTickfileSpecifiers = New TickFileSpecifiers
mTickfileSpecifiers.Add Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TickFileSpecifiers(ByVal Value As TickFileSpecifiers)
Const ProcName As String = "TickfileSpecifiers"
On Error GoTo Err

Assert Not mReplayInProgress, "Tickfile is already replaying"

Set mTickfileSpecifiers = Value

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

Public Sub ClearTickfileSpecifiers()
Const ProcName As String = "ClearTickfileSpecifiers"
On Error GoTo Err

mTickfileSpecifiers.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pReplaySpeed As Long, _
                ByVal pTimestampAdjustmentStart As Double, _
                ByVal pTimestampAdjustmentEnd As Double, _
                ByVal pReplayProgressEventInterval As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTickfileStore = pTickfileStore
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore
mReplaySpeed = pReplaySpeed
mTimestampAdjustmentStart = pTimestampAdjustmentStart
mTimestampAdjustmentEnd = pTimestampAdjustmentEnd

mReplayProgressEventInterval = pReplayProgressEventInterval

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "NotifyEvent"
On Error GoTo Err

Set ev.Source = Me
RaiseEvent NotifyEvent(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyProgress(ByVal pTimestamp As Date, ByVal pEventsPlayed As Long, ByVal pPercentComplete As Long)
Const ProcName As String = "NotifyProgress"
On Error GoTo Err

If pPercentComplete > 100 Then pPercentComplete = 100
RaiseEvent ReplayProgress(pTimestamp, pEventsPlayed, pPercentComplete)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyReady(ByVal pPlayer As TickfilePlayer)
Const ProcName As String = "NotifyReady"
On Error GoTo Err

mControllerTask.NotifyReady pPlayer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifySkipTickfile(ByVal pPlayer As TickfilePlayer, ByRef ev As TickfileEventData, ByVal pEventsPlayed As Long)
Const ProcName As String = "NotifySkipTickfile"
On Error GoTo Err

mControllerTask.NotifySkip pPlayer
fireTickfileCompleted ev, pEventsPlayed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTickfileCompleted(ByVal pPlayer As TickfilePlayer, ByRef ev As TickfileEventData, ByVal pEventsPlayed As Long)
Const ProcName As String = "NotifyReplayComplete"
On Error GoTo Err

mControllerTask.NotifyTickfileCompleted pPlayer
fireTickfileCompleted ev, pEventsPlayed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTickStreamStarted(ByRef ev As TickfileEventData, ByVal pClock As Clock)
Const ProcName As String = "NotifyReady"
On Error GoTo Err

RaiseEvent TickStreamStarted(ev, pClock)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PauseReplay()
Const ProcName As String = "PauseReplay"
On Error GoTo Err

Assert mReplayInProgress, "Tickfile is not replaying"

mControllerTask.PauseReplay

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ResumeReplay()
Const ProcName As String = "ResumeReplay"
On Error GoTo Err

Assert mReplayInProgress, "Replay has not been started"
mControllerTask.ResumeReplay

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function SetupTickfilePlayer( _
                ByRef pTickfileSpecifier As TickfileSpecifier, _
                ByVal pTickScheduler As TickScheduler) As TickfilePlayer
Const ProcName As String = "SetupTickfilePlayer"
On Error GoTo Err

Dim lStreamId As Long
lStreamId = getNextStreamId

Dim lTickfilePlayer As TickfilePlayer
Set lTickfilePlayer = New TickfilePlayer
lTickfilePlayer.Initialise Me, _
                    pTickScheduler, _
                    mPrimaryContractStore, _
                    mSecondaryContractStore, _
                    lStreamId, _
                    mReplaySpeed, _
                    mTimestampAdjustmentStart, _
                    mTimestampAdjustmentEnd

Dim lTickfileReader As ITickfileReader
Set lTickfileReader = mTickfileStore.CreateTickfileReader(lTickfilePlayer, lTickfilePlayer, lStreamId)

If lTickfileReader Is Nothing Then
ElseIf lTickfilePlayer.PrepareTickfile(lTickfileReader, pTickfileSpecifier) Then
    Set SetupTickfilePlayer = lTickfilePlayer
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SkipTickfile()
Const ProcName As String = "SkipTickfile"
On Error GoTo Err

Assert mReplayInProgress, "Replay has not been started"
Assert Not mMerge, "Not permitted for merged replay"

Dim lPlayer As TickfilePlayer
Set lPlayer = mControllerTask.CurrentPlayer

If lPlayer Is Nothing Then Exit Sub

Dim ev As TickfileEventData
ev = gGetTickfileEventData(Me, mControllerTask.CurrentTickfileSpec, lPlayer)

Dim lEventsPlayed As Long
lEventsPlayed = lPlayer.ReplayProgressTicks

mControllerTask.NotifySkip lPlayer
fireTickfileCompleted ev, lEventsPlayed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function StartReplay(Optional ByVal pMerge As Boolean = False) As TaskController
Const ProcName As String = "StartReplay"
On Error GoTo Err

Assert Not mReplayInProgress, "Replay has already been started"
    
mReplayInProgress = True
mMerge = pMerge

If mMerge Then
    Set mControllerTask = New MergedClrTask
Else
    Set mControllerTask = New SequentialClrTask
End If

mControllerTask.Initialise Me, mTickfileSpecifiers, mReplaySpeed, mReplayProgressEventInterval
Set StartReplay = StartTask(mControllerTask, PriorityNormal)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub StopReplay()
Const ProcName As String = "StopReplay"
On Error GoTo Err

If Not mReplayInProgress Then Exit Sub
mReplayInProgress = False

mControllerTask.StopReplay

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireTickfileCompleted(ByRef ev As TickfileEventData, ByVal pEventsPlayed As Long)
Const ProcName As String = "fireTickfileCompleted"
On Error GoTo Err

RaiseEvent TickfileCompleted(ev, pEventsPlayed)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getNextStreamId() As Long
Static sNextStreamId As Long
getNextStreamId = sNextStreamId
sNextStreamId = sNextStreamId + 1
End Function

Friend Function QueryPlay( _
                ByRef pTickfileSpec As ITickfileSpecifier, _
                ByVal pPlayer As TickfilePlayer) As ReplayContinueModes
Const ProcName As String = "playOrNot"
On Error GoTo Err

Dim lContinueMode As ReplayContinueModes

RaiseEvent QueryReplayNextTickfile(gGetTickfileEventData(Me, pTickfileSpec, pPlayer), lContinueMode)
QueryPlay = lContinueMode

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


