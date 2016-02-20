VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITaskCompletionListener

Private mTickfileStore                          As ITickfileStore
Private mTickfileSpecs                          As TickFileSpecifiers

Private mTestCompleted                          As Boolean

Private mEt                                     As New ElapsedTimer

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

Set mTickfileStore = CreateTickfileStore(TickfileReadOnly, TickfilePath, TickfileGranularityWeek)
Set mTickfileSpecs = New TickFileSpecifiers

End Sub

Public Sub TearDown()
LogMessage "Test Completed -----------------------------------------------"
End Sub

'@================================================================================
' Interface Members
'@================================================================================

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

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 1, "NotifyEvent should be invoked once"
Assert.IsTrue tfmev.NotifyEventEv.EventCode = TickfileEventCodes.TickfileEventFileDoesNotExist, "NotifyEvent has invalid event code"
Assert.IsTrue tfmev.ReplayProgressCount = 0, "ReplayProgress should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 0, "TickfileCompleted should not be invoked"
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
Assert.IsTrue tfmev.TickfileCompletedCount = 0, "TickfileCompleted should not be invoked"
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
Assert.IsTrue tfmev.NotifyEventEv.EventCode = TickfileEventCodes.TickfileEventFileIsEmpty, "NotifyEvent has invalid event code"
Assert.IsTrue tfmev.ReplayProgressCount = 0, "ReplayProgress should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should not be invoked"
End Sub

Public Sub ReplayWithValidFileTradeBuildv5()
mTickfileSpecs.Add CreateTickfileSpecWithValidTradeBuildv5File

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, , , 0), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
Assert.AreEqual 105, tfmev.NumberOfReceivedTicks, "Incorrect number of ticks received"
End Sub

Public Sub ReplayWithValidFileCrescendov1NoContractStore()
mTickfileSpecs.Add CreateTickfileSpecWithValidCrescendov1File

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, , , 0), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 1, "NotifyEvent should be invoked once"
Assert.IsTrue tfmev.NotifyEventEv.EventCode = TickfileEventCodes.TickfileEventNoContractDetails, "NotifyEvent has invalid event code"
Assert.IsTrue tfmev.ReplayProgressCount = 0, "ReplayProgress should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 0, "TickfileCompleted should not be invoked"
End Sub

Public Sub ReplayWithValidFileCrescendov1()
mTickfileSpecs.Add CreateTickfileSpecWithValidCrescendov1File

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, New MockContractStore, , 0), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Dim lStartTime As Date
lStartTime = Now
Do While Not mTestCompleted
    Wait 50
    'If (Now - lStartTime) > 2 * OneSecond Then Exit Do
Loop

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
Assert.AreEqual 44110, tfmev.NumberOfReceivedTicks, "Incorrect number of ticks received"
End Sub

Public Sub ReplayWithVeryLargeValidFileTradeBuildv5()
mTickfileSpecs.Add CreateTickfileSpecWithVeryLargeValidTradeBuildv5File

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, , , 0), mTickfileSpecs

mEt.StartTiming
tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 5
Loop Until mTestCompleted

Dim t As Single
t = mEt.ElapsedTimeMicroseconds

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
Assert.AreEqual 501623, tfmev.NumberOfReceivedTicks, "Incorrect number of ticks received"

LogMessage "Replayed " & tfmev.TickfileCompletedEventsPlayed & " events in " & Format(t / 1000000#, "0.000 seconds")
LogMessage CStr(Int(tfmev.TickfileCompletedEventsPlayed / (t / 1000000#))) & " ticks per second"
LogMessage CStr(Int(t / tfmev.TickfileCompletedEventsPlayed)) & " microseconds per tick"
End Sub

Public Sub ReplayWithValidFileCrescendov1WithFiltering()
Dim lTfs As TickfileSpecifier

Set lTfs = CreateTickfileSpecWithValidCrescendov1File
lTfs.FromDate = CDate("06/06/2003 14:51:44")
lTfs.ToDate = CDate("06/06/2003 14:51:56")
mTickfileSpecs.Add lTfs

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, New MockContractStore, , 0), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
Assert.AreEqual 38, tfmev.NumberOfReceivedTicks, "Incorrect number of ticks received"
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function CreateTickfileSpecWithEmptyFile() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.TickfileFormatID = "Valid format"
tickfileSpec.FileName = TickfilePath & "\Empty file.tck"
Set CreateTickfileSpecWithEmptyFile = tickfileSpec
End Function

Private Function CreateTickfileSpecWithInvalidFile() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.TickfileFormatID = "Valid format"
tickfileSpec.FileName = TickfilePath & "\Invalid file.tck"
Set CreateTickfileSpecWithInvalidFile = tickfileSpec
End Function

Private Function CreateTickfileSpecWithInvalidFilename() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.TickfileFormatID = "Valid format"
tickfileSpec.FileName = "Invalid filename"
Set CreateTickfileSpecWithInvalidFilename = tickfileSpec
End Function

Private Function CreateTickfileSpecWithValidCrescendov1File() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.FileName = TickfilePath & "\Z20030620030606075945.csv"
Set CreateTickfileSpecWithValidCrescendov1File = tickfileSpec
End Function

Private Function CreateTickfileSpecWithValidTradeBuildv5File() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.FileName = TickfilePath & "\ZH3-Session-20130308.tck"
Set CreateTickfileSpecWithValidTradeBuildv5File = tickfileSpec
End Function

Private Function CreateTickfileSpecWithVeryLargeValidTradeBuildv5File() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.FileName = TickfilePath & "\ZM3-Week-20130408.tck"
Set CreateTickfileSpecWithVeryLargeValidTradeBuildv5File = tickfileSpec
End Function


