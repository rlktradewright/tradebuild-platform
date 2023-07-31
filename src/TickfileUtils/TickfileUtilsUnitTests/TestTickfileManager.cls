VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTickfileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IGenericTickListener
Implements ITaskCompletionListener

Private mTickfileStore                          As ITickfileStore
Private mTickfileStoreAsync                     As ITickfileStore
Private mTickfileSpecs                          As TickFileSpecifiers

Private mTestCompleted                          As Boolean

Private mTickNumber                             As Long

'@================================================================================
' Setup Functions
'@================================================================================

Public Sub FixtureSetup()
End Sub

Public Sub FixtureTeardown()
End Sub

Public Sub Setup()
mTestCompleted = False
ReDim ReceivedTicks(15) As TickInfo
NumberOfReceivedTicks = 0

Set mTickfileStore = New MockTickfileStore

Set mTickfileStoreAsync = New MockTickfileStore
Dim tfStore As MockTickfileStore
Set tfStore = mTickfileStoreAsync
tfStore.Async = True

Set mTickfileSpecs = New TickFileSpecifiers

mTickNumber = 0
End Sub

Public Sub TearDown()
LogMessage "Test Completed -----------------------------------------------"
End Sub

'@================================================================================
' Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
End Sub

Private Sub IGenericTickListener_NotifyTick(ByRef ev As GenericTickEventData)
mTickNumber = mTickNumber + 1
gAddReceivedTick ev.Tick
'LogMessage FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  Tick " & mTickNumber & ": " & GenericTickToString(ev.Tick)
End Sub

Private Sub ITaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
mTestCompleted = True
End Sub

'@================================================================================
' Test Methods
'@================================================================================

Public Sub ReplayWithNonExistentFile()
mTickfileSpecs.Add CreateTickfileSpecWithInvalidFilename

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore), mTickfileSpecs

tfmev.ReplayController.StartReplay().AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 1, "NotifyEvent should be invoked once"
Assert.IsTrue tfmev.NotifyEventEv.EventCode = TickfileEventCodes.TickfileEventFileDoesNotExist, "NotifyEvent has invalid event code"
Assert.IsTrue tfmev.ReplayProgressCount = 0, "ReplayProgress should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 0, "TickfileCompleted should be invoked once"
End Sub

Public Sub ReplayWithEmptyFile()
mTickfileSpecs.Add CreateTickfileSpecWithEmptyFile

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 1, "NotifyEvent should be invoked once"
Assert.IsTrue tfmev.NotifyEventEv.EventCode = TickfileEventCodes.TickfileEventNoTickData, "NotifyEvent has invalid event code"
Assert.IsTrue tfmev.ReplayProgressCount = 0, "ReplayProgress should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should not be invoked"
Assert.AreEqual 0, tfmev.NumberOfReceivedTicks, "Incorrect number of ticks received"
End Sub

Public Sub ReplayWithInvalidFile()
mTickfileSpecs.Add CreateTickfileSpecWithInvalidFile

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 1, "NotifyEvent should be invoked once"
Assert.IsTrue tfmev.NotifyEventEv.EventCode = TickfileEventCodes.TickfileEventFileIsInvalid, "NotifyEvent has invalid event code"
Assert.IsTrue tfmev.ReplayProgressCount = 0, "ReplayProgress should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 0, "TickfileCompleted should be invoked once"
End Sub

Public Sub ReplayWithValidFile()
mTickfileSpecs.Add CreateTickfileSpecWithValidFileA

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.ReplayProgressCount > 0, "ReplayProgress should be invoked at least once"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
gCheckReceivedTicks TicksA
End Sub

Public Sub ReplayWithValidFileAndPauseResume()
mTickfileSpecs.Add CreateTickfileSpecWithValidFileA

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Wait 1250
tfmev.ReplayController.PauseReplay
Wait 2000
tfmev.ReplayController.ResumeReplay

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
gCheckReceivedTicks TicksA, , 1250, 2000
End Sub

'Public Sub ReplayWithTwoValidFiles()
'mTickfileSpecs.Add CreateTickfileSpecWithValidFileA
'mTickfileSpecs.Add CreateTickfileSpecWithValidFileB
'
'Dim tfmev As New TickfileMgrEV
'tfmev.Initialise CreateTickfileReplayController(mTickfileStore), mTickfileSpecs
'
'tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me
'
'Do
'    Wait 50
'Loop Until mTestCompleted
'
'Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
'Assert.IsTrue tfmev.QueryReplayNextTickfileCount = 2, "QueryReplayNextTickfile should be invoked twice"
'Assert.IsTrue tfmev.ReplayProgressCount >= 2, "ReplayProgress should be invoked at least twice"
'Assert.IsTrue tfmev.TickfileCompletedCount = 2, "TickfileCompleted should be invoked twice"
'Assert.IsTrue tfmev.TickStreamStartedCount = 2, "TickStreamStarted should should be invoked twice"
'Assert.IsTrue tfmev.TickStreamStartedEv.TickStream.StreamId = 1, "StreamId should be 1"
'Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
'Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
'Assert.AreSame GetTimeZone("Central Standard Time"), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be CST"
'gCheckReceivedTicksTwoLots TicksA, TicksB
'End Sub

'Public Sub ReplayWithTwoValidFilesAndSkip()
'mTickfileSpecs.Add CreateTickfileSpecWithValidFileA
'mTickfileSpecs.Add CreateTickfileSpecWithValidFileB
'
'Dim tfmev As New TickfileMgrEV
'tfmev.Initialise CreateTickfileReplayController(mTickfileStore), mTickfileSpecs
'
'tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me
'
'Wait 1750
'
'tfmev.ReplayController.SkipTickfile
'
'Do
'    Wait 50
'Loop Until mTestCompleted
'
'Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
'Assert.IsTrue tfmev.QueryReplayNextTickfileCount = 2, "QueryReplayNextTickfile should be invoked twice"
'Assert.IsTrue tfmev.ReplayProgressCount >= 2, "ReplayProgress should be invoked at least twice"
'Assert.IsTrue tfmev.TickfileCompletedCount = 2, "TickfileCompleted should be invoked twice"
'Assert.IsTrue tfmev.TickStreamStartedCount = 2, "TickStreamStarted should should be invoked twice"
'Assert.IsTrue tfmev.TickStreamStartedEv.TickStream.StreamId = 1, "StreamId should be 1"
'Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
'Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
'Assert.AreSame GetTimeZone("Central Standard Time"), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be CST"
'gCheckReceivedTicksTwoLots TicksA, TicksB, 2
'End Sub

Public Sub ReplayWithValidFileAsync()
mTickfileSpecs.Add CreateTickfileSpecWithValidFileA

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStoreAsync), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.ReplayProgressCount > 0, "ReplayProgress should be invoked at least once"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
gCheckReceivedTicks TicksA
End Sub

'Public Sub ReplayWithTwoValidFilesAsync()
'mTickfileSpecs.Add CreateTickfileSpecWithValidFileA
'mTickfileSpecs.Add CreateTickfileSpecWithValidFileB
'
'Dim tfmev As New TickfileMgrEV
'tfmev.Initialise CreateTickfileReplayController(mTickfileStoreAsync)
'
'tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me
'
'Do
'    Wait 50
'Loop Until mTestCompleted
'
'Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
'Assert.IsTrue tfmev.QueryReplayNextTickfileCount = 2, "QueryReplayNextTickfile should be invoked twice"
'Assert.IsTrue tfmev.ReplayProgressCount >= 2, "ReplayProgress should be invoked at least twice"
'Assert.IsTrue tfmev.TickfileCompletedCount = 2, "TickfileCompleted should be invoked twice"
'Assert.IsTrue tfmev.TickStreamStartedCount = 2, "TickStreamStarted should should be invoked twice"
'Assert.IsTrue tfmev.TickStreamStartedEv.TickStream.StreamId = 1, "StreamId should be 1"
'Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
'Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
'Assert.AreSame GetTimeZone("Central Standard Time"), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be CST"
'gCheckReceivedTicksTwoLots TicksA, TicksB
'End Sub

Public Sub ConcurrentReplayWithTwoValidFiles()
mTickfileSpecs.Add CreateTickfileSpecWithValidFileA
mTickfileSpecs.Add CreateTickfileSpecWithValidFileB

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore), mTickfileSpecs

tfmev.ReplayController.StartReplay().AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.ReplayProgressCount >= 1, "ReplayProgress should be invoked at least once"
Assert.IsTrue tfmev.TickfileCompletedCount = 2, "TickfileCompleted should be invoked twice"
gCheckReceivedTicksConcurrent
End Sub

Public Sub ConcurrentReplayWithTwoValidFilesAsync()
mTickfileSpecs.Add CreateTickfileSpecWithValidFileA
mTickfileSpecs.Add CreateTickfileSpecWithValidFileB

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStoreAsync), mTickfileSpecs

tfmev.ReplayController.StartReplay().AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.ReplayProgressCount >= 1, "ReplayProgress should be invoked at least once"
Assert.IsTrue tfmev.TickfileCompletedCount = 2, "TickfileCompleted should be invoked twice"
gCheckReceivedTicksConcurrent
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function CreateTickfileSpecWithEmptyFile() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.TickfileFormatID = "Valid format"
tickfileSpec.FileName = "Empty file"
Set CreateTickfileSpecWithEmptyFile = tickfileSpec
End Function

Private Function CreateTickfileSpecWithInvalidFile() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.TickfileFormatID = "Valid format"
tickfileSpec.FileName = "Invalid file"
Set CreateTickfileSpecWithInvalidFile = tickfileSpec
End Function

'Private Function CreateTickfileSpecWithInvalidFormatId() As TickfileSpecifier
'Dim tickfileSpec As New TickfileSpecifier
'tickfileSpec.TickfileFormatID = "Invalid format"
'Set CreateTickfileSpecWithInvalidFormatId = tickfileSpec
'End Function

Private Function CreateTickfileSpecWithInvalidFilename() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.TickfileFormatID = "Valid format"
tickfileSpec.FileName = "Invalid filename"
Set CreateTickfileSpecWithInvalidFilename = tickfileSpec
End Function

Private Function CreateTickfileSpecWithValidFileA() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.TickfileFormatID = "Valid format"
tickfileSpec.FileName = "Valid A"
Set CreateTickfileSpecWithValidFileA = tickfileSpec
End Function

Private Function CreateTickfileSpecWithValidFileB() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.TickfileFormatID = "Valid format"
tickfileSpec.FileName = "Valid B"
Set CreateTickfileSpecWithValidFileB = tickfileSpec
End Function



