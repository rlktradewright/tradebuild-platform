VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTickfileReaderDB"
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

Set mTickfileStore = CreateTradingDBClient(DbSQLServer2005, "Sappy\Live", "TradingDB", pUseSynchronousReads:=True, pUseSynchronousWrites:=True).TickfileStore
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

Public Sub ReplayWithEmptyDatabaseTickfile()
Dim lTickfileSpec As New TickfileSpecifier
lTickfileSpec.Contract = gCreateContractFromLocalSymbol("ZM03")
lTickfileSpec.FromDate = CDate("2014/09/09")
lTickfileSpec.ToDate = CDate("2014/09/10")

mTickfileSpecs.Add lTickfileSpec

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, , , 0), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue tfmev.NotifyEventCount = 1, "NotifyEvent should be invoked once"
Assert.IsTrue tfmev.NotifyEventEv.EventCode = TickfileEventCodes.TickfileEventNoTickData, "NotifyEvent has invalid event code"
Assert.IsTrue tfmev.ReplayProgressCount = 0, "ReplayProgress should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
End Sub

Public Sub ReplayWithOldDatabaseTickfile()
Dim lTickfileSpec As New TickfileSpecifier
lTickfileSpec.Contract = gCreateContractFromLocalSymbol("ZM03")
lTickfileSpec.FromDate = CDate("2003/06/19")
lTickfileSpec.ToDate = CDate("2003/06/20")

mTickfileSpecs.Add lTickfileSpec

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, , , 0), mTickfileSpecs

mEt.StartTiming
tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTestCompleted

Dim t As Single
t = mEt.ElapsedTimeMicroseconds
LogMessage "Replayed " & tfmev.TickfileCompletedEventsPlayed & " events in " & Format(t / 1000000#, "0.000 seconds")
LogMessage CStr(Int(tfmev.TickfileCompletedEventsPlayed / (t / 1000000#))) & " ticks per second"
LogMessage CStr(Int(t / tfmev.TickfileCompletedEventsPlayed)) & " microseconds per tick"

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
Assert.AreEqual 44382, tfmev.NumberOfReceivedTicks, "Incorrect number of ticks received"

End Sub

Public Sub ReplayWithDatabaseTickfile()
Dim lTickfileSpec As New TickfileSpecifier
lTickfileSpec.Contract = gCreateContractFromLocalSymbol("ZH13")
lTickfileSpec.FromDate = CDate("2013/03/08 18:19:39")
lTickfileSpec.ToDate = CDate("2013/03/08 18:19:58")

mTickfileSpecs.Add lTickfileSpec

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
Assert.AreEqual 19, tfmev.NumberOfReceivedTicks, "Incorrect number of ticks received"
End Sub

Public Sub ReplayWithVeryLargeValidDatabaseTickfile()
Dim lTickfileSpec As New TickfileSpecifier
lTickfileSpec.Contract = gCreateContractFromLocalSymbol("ZM13")
lTickfileSpec.FromDate = CDate("2013/04/08")
lTickfileSpec.ToDate = CDate("2013/04/13")

mTickfileSpecs.Add lTickfileSpec

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, , , 0), mTickfileSpecs

mEt.StartTiming
tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 5
Loop Until mTestCompleted

Dim t As Single
t = mEt.ElapsedTimeMicroseconds
LogMessage "Replayed " & tfmev.TickfileCompletedEventsPlayed & " events in " & Format(t / 1000000#, "0.000 seconds")
LogMessage CStr(Int(tfmev.TickfileCompletedEventsPlayed / (t / 1000000#))) & " ticks per second"
LogMessage CStr(Int(t / tfmev.TickfileCompletedEventsPlayed)) & " microseconds per tick"

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
Assert.AreEqual 505173, tfmev.NumberOfReceivedTicks, "Incorrect number of ticks received"
End Sub

Public Sub ReplayWithVeryLargeValidDatabaseTickfileSessionOnly()
Dim lTickfileSpec As New TickfileSpecifier
lTickfileSpec.Contract = gCreateContractFromLocalSymbol("ZU14")
lTickfileSpec.FromDate = CDate("2014/09/01")
lTickfileSpec.ToDate = CDate("2014/09/05")
lTickfileSpec.EntireSession = True

mTickfileSpecs.Add lTickfileSpec

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, , , 0), mTickfileSpecs

mEt.StartTiming
tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 5
Loop Until mTestCompleted

Dim t As Single
t = mEt.ElapsedTimeMicroseconds
LogMessage "Replayed " & tfmev.TickfileCompletedEventsPlayed & " events in " & Format(t / 1000000#, "0.000 seconds")
LogMessage CStr(Int(tfmev.TickfileCompletedEventsPlayed / (t / 1000000#))) & " ticks per second"
LogMessage CStr(Int(t / tfmev.TickfileCompletedEventsPlayed)) & " microseconds per tick"

Assert.IsTrue tfmev.NotifyEventCount = 0, "NotifyEvent should not be invoked"
Assert.IsTrue tfmev.TickfileCompletedCount = 1, "TickfileCompleted should be invoked once"
Assert.IsNotNothing tfmev.TickStreamStartedClock, "Clock should not be nothing"
Assert.IsTrue tfmev.TickStreamStartedClock.IsSimulated, "Clock should be simulated"
Assert.AreSame GetTimeZone(""), tfmev.TickStreamStartedClock.TimeZone, "Timezone should be local"
Assert.AreEqual 452426, tfmev.NumberOfReceivedTicks, "Incorrect number of ticks received"
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
tickfileSpec.FileName = TickfilePath & "\ZM13-Week-20130408.tck"
Set CreateTickfileSpecWithVeryLargeValidTradeBuildv5File = tickfileSpec
End Function



