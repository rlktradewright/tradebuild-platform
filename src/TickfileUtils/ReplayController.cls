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

'@================================================================================
' Events
'@================================================================================

Event NotifyEvent(ev As NotificationEventData)

Event ReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Long)

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

Private mTickScheduler                              As TickScheduler

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

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

Dim lPlayer As TickfilePlayer
Dim i As Long
For i = 0 To UBound(mPlayers)
    Set lPlayer = mPlayers(i)
    lPlayer.ReplaySpeed = mReplaySpeed
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

Set TickStream = mPlayers(pIndex).TickStream

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
                ByVal pTickfileSpecifiers As TickFileSpecifiers, _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pReplaySpeed As Long, _
                ByVal pTimestampAdjustmentStart As Double, _
                ByVal pTimestampAdjustmentEnd As Double, _
                ByVal pReplayProgressEventInterval As Long)
Const ProcName As String = "Intialise"
On Error GoTo Err

Set mTickfileSpecifiers = pTickfileSpecifiers.Clone
Set mTickfileStore = pTickfileStore
Set mPrimaryContractStore = pPrimaryContractStore
Set pSecondaryContractStore = pSecondaryContractStore
mReplaySpeed = pReplaySpeed
mTimestampAdjustmentStart = pTimestampAdjustmentStart
mTimestampAdjustmentEnd = pTimestampAdjustmentEnd

mReplayProgressEventInterval = pReplayProgressEventInterval

ReDim mPlayers(mTickfileSpecifiers.Count - 1) As TickfilePlayer

Set mTickScheduler = New TickScheduler
mTickScheduler.Initialise Me, mReplaySpeed, mReplayProgressEventInterval, IIf(mTickfileSpecifiers.Count = 1, True, False)

Dim i As Long
For i = 0 To UBound(mPlayers)
    Set mPlayers(i) = SetupTickfilePlayer(mTickfileSpecifiers(i + 1), mTickScheduler)
Next

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

fireProgress pTimestamp, pEventsPlayed, pPercentComplete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyReady(ByVal pPlayer As TickfilePlayer)
Const ProcName As String = "NotifyReady"
On Error GoTo Err

mNumberOfPlayersAwaitingReady = mNumberOfPlayersAwaitingReady - 1

mTickScheduler.AddPlayer pPlayer
pPlayer.StartReplay

If mNumberOfPlayersAwaitingReady = 0 And mReplayInProgress Then mTickScheduler.StartScheduling

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTickfileCompleted(ByVal pPlayer As TickfilePlayer, ByRef ev As TickfileEventData, ByVal pEventsPlayed As Long)
Const ProcName As String = "NotifyReplayComplete"
On Error GoTo Err

fireTickfileCompleted ev, pEventsPlayed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PauseReplay()
Const ProcName As String = "PauseReplay"
On Error GoTo Err

Assert mReplayInProgress, "Replay has not yet been started"

mPaused = True

Dim lPlayer As TickfilePlayer
Dim i As Long
For i = 0 To UBound(mPlayers)
    Set lPlayer = mPlayers(i)
    lPlayer.PauseReplay
Next

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

Dim lPlayer As TickfilePlayer
Dim i As Long
For i = 0 To UBound(mPlayers)
    Set lPlayer = mPlayers(i)
    lPlayer.ResumeReplay
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function StartReplay() As TaskController
Const ProcName As String = "StartReplay"
On Error GoTo Err

Assert Not mReplayInProgress, "Replay has already been started"
    
mReplayInProgress = True

If mNumberOfPlayersAwaitingReady = 0 Then mTickScheduler.StartScheduling

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub StopReplay()
Const ProcName As String = "StopReplay"
On Error GoTo Err

Assert mReplayInProgress, "Replay has not yet been started"
mReplayInProgress = False

Dim lPlayer As TickfilePlayer
Dim i As Long
For i = 0 To UBound(mPlayers)
    Set lPlayer = mPlayers(i)
    lPlayer.StopReplay
Next

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

Dim lListener As ITickfileReplayListener
For Each lListener In mListeners.CurrentListeners
    lListener.NotifyEvent ev
Next

RaiseEvent NotifyEvent(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireProgress(ByVal pTimestamp As Date, ByVal pEventsPlayed As Long, ByVal pPercentComplete As Long)
Const ProcName As String = "fireProgress"
On Error GoTo Err

Dim lListener As ITickfileReplayListener
For Each lListener In mListeners.CurrentListeners
    lListener.ReplayProgress pTimestamp, pEventsPlayed, pPercentComplete
Next

RaiseEvent ReplayProgress(pTimestamp, pEventsPlayed, pPercentComplete)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireTickfileCompleted(ByRef ev As TickfileEventData, ByVal pEventsPlayed As Long)
Const ProcName As String = "fireTickfileCompleted"
On Error GoTo Err

Dim lListener As ITickfileReplayListener
For Each lListener In mListeners.CurrentListeners
    lListener.TickfileCompleted ev, pEventsPlayed
Next

RaiseEvent TickfileCompleted(ev, pEventsPlayed)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function SetupTickfilePlayer( _
                ByRef pTickfileSpecifier As TickfileSpecifier, _
                ByVal pTickScheduler As TickScheduler) As TickfilePlayer
Const ProcName As String = "SetupTickfilePlayer"
On Error GoTo Err

Dim lStreamId As Long
lStreamId = gGetNextStreamId

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

Assert Not lTickfileReader Is Nothing, "Failed to create TickfileReader"

If lTickfilePlayer.PrepareTickfile(lTickfileReader, pTickfileSpecifier) Then
    Set SetupTickfilePlayer = lTickfilePlayer
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



