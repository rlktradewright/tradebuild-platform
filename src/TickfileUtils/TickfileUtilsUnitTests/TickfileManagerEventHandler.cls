VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileMgrEV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IGenericTickListener
Implements IStateChangeListener
Implements ITickfileReplayListener

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

Private Const ModuleName                            As String = "TickfileMgrEV"

'@================================================================================
' Member variables
'@================================================================================

Public ReplayController                             As ReplayController
Attribute ReplayController.VB_VarHelpID = -1

Private mNotifyEventEv                              As NotificationEventData
Public ReplayProgressTickfileTimestamp              As Date
Public ReplayProgressEventsPlayed                   As Long
Public ReplayProgressPercentComplete                As Single
Private mTickfileCompletedEv                        As TickfileEventData
Public TickfileCompletedEventsPlayed                As Long
Public TickStreamId                                 As Long
Public TickStreamStartedClock                       As Clock

Public NotifyEventCount                             As Long
Public ReplayProgressCount                          As Long
Public TickfileCompletedCount                       As Long
Public TickStreamStartedCount                       As Long

Public NumberOfReceivedTicks                        As Long

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
LogMessage "TickfileMgrEV created"
End Sub

'@================================================================================
' IGenericTickListener Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As TickUtils27.GenericTickEventData)
End Sub

Private Sub IGenericTickListener_NotifyTick(ByRef ev As GenericTickEventData)
NumberOfReceivedTicks = NumberOfReceivedTicks + 1
gAddReceivedTick ev.Tick
'LogMessage "  Tick " & NumberOfReceivedTicks & ": " & GenericTickToString(ev.Tick)
End Sub

'@================================================================================
' IStateChangeListener Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
If ev.State = TickStreamStates.TickStreamStateRunning Then
    TickStreamStartedCount = TickStreamStartedCount + 1
    Dim lStream As ITickStream
    Set lStream = ev.Source
    TickStreamId = lStream.StreamId
    LogMessage "TickStream started: count=" & TickStreamStartedCount
End If
End Sub

'@================================================================================
' ITickfileReplayListener Members
'@================================================================================

Private Sub ITickfileReplayListener_NotifyEvent(ev As NotificationEventData)
NotifyEventCount = NotifyEventCount + 1
mNotifyEventEv = ev
LogMessage "ITickfileReplayListener_NotifyEvent: count=" & NotifyEventCount & "; message=" & ev.EventMessage
End Sub

Private Sub ITickfileReplayListener_ReplayProgress(ByVal pTickfileTimestamp As Date, ByVal pEventsPlayed As Long, ByVal pPercentComplete As Long)
ReplayProgressCount = ReplayProgressCount + 1
ReplayProgressEventsPlayed = pEventsPlayed
ReplayProgressPercentComplete = pPercentComplete
ReplayProgressTickfileTimestamp = pTickfileTimestamp
End Sub

Private Sub ITickfileReplayListener_TickfileCompleted(ev As TickfileEventData, ByVal pEventsPlayed As Long)
TickfileCompletedCount = TickfileCompletedCount + 1
mTickfileCompletedEv = ev
TickfileCompletedEventsPlayed = pEventsPlayed
LogMessage "ITickfileReplayListener_TickfileCompleted: count=" & TickfileCompletedCount
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
If Not ev.Future.IsAvailable Then Exit Sub
Set TickStreamStartedClock = ev.Future.Value
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get NotifyEventEv() As NotificationEventData
NotifyEventEv = mNotifyEventEv
End Property

Public Property Get TickfileCompletedEv() As TickfileEventData
TickfileCompletedEv = mTickfileCompletedEv
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pReplayController As ReplayController, _
                    ByVal pTickfileSpecs As TickFileSpecifiers)
Set ReplayController = pReplayController
ReplayController.AddTickfileReplayListener Me
ReplayController.TickFileSpecifiers = pTickfileSpecs

Dim i As Long
Dim lTickStream As ITickStream
For i = 0 To pTickfileSpecs.Count - 1
    Set lTickStream = ReplayController.TickStream(i)
    If Not lTickStream Is Nothing Then
        lTickStream.AddGenericTickListener Me
        lTickStream.AddStateChangeListener Me
        mFutureWaiter.Add lTickStream.ClockFuture
        lTickStream.Start
    End If
Next

End Sub

'@================================================================================
' Helper Functions
'@================================================================================




