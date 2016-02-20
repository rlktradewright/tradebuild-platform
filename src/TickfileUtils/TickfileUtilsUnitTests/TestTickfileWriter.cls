VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTickfileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IGenericTickListener
Implements ITickfileOutputMonitor
Implements ITaskCompletionListener

Private Const TickfileAPath                     As String = TickfilePath & "\ZH3\ZH3-Week-20130218.tck"

Private mTickfileStore                          As ITickfileStore
Private mTickfileSpecs                          As TickFileSpecifiers

Private mTickNumber                             As Long

Private mNotifyEventEv                          As NotificationEventData

Private mNotifyEventCount                       As Long
Private mNotifyReadyCount                       As Long
Private mNotifyOutputTickfileClosedCount        As Long
Private mOutputTickfileName                     As String

Private mContractFuture                         As IFuture

Private mTaskCompleted                          As Boolean

'@================================================================================
' Setup Functions
'@================================================================================

Public Sub FixtureSetup()
End Sub

Public Sub FixtureTeardown()
End Sub

Public Sub Setup()
mTaskCompleted = False
ReDim ReceivedTicks(15) As TickInfo
NumberOfReceivedTicks = 0

Set mTickfileStore = CreateTickfileStore(TickfileWriteOnly, TickfilePath, TickfileGranularityWeek)
Set mTickfileSpecs = New TickFileSpecifiers

mTickNumber = 0

mNotifyEventCount = 0
mNotifyEventEv.EventCode = 0
mNotifyEventEv.EventMessage = ""
Set mNotifyEventEv.Source = Nothing
mNotifyReadyCount = 0
mNotifyOutputTickfileClosedCount = 0
mOutputTickfileName = ""

Set mContractFuture = CreateFuture(gCreateContractFromLocalSymbol("ZH3"))

End Sub

Public Sub TearDown()
LogMessage "Test Completed -----------------------------------------------"
End Sub

'@================================================================================
' Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As TickUtils27.GenericTickEventData)
End Sub

Private Sub IGenericTickListener_NotifyTick(ByRef ev As GenericTickEventData)
mTickNumber = mTickNumber + 1
gAddReceivedTick ev.Tick
LogMessage FormatTimestamp(GetTimestamp, TimestampDateAndTimeISO8601) & "  Tick " & mTickNumber & ": " & GenericTickToString(ev.Tick)
End Sub

Private Sub ITickfileOutputMonitor_NotifyEvent(ev As NotificationEventData)
mNotifyEventCount = mNotifyEventCount + 1
End Sub

Private Sub ITickfileOutputMonitor_NotifyNotReady(ByVal pSource As Object)

End Sub

Private Sub ITickfileOutputMonitor_NotifyOutputFileClosed(ByVal pSource As Object)
mNotifyOutputTickfileClosedCount = mNotifyOutputTickfileClosedCount + 1
End Sub

Private Sub ITickfileOutputMonitor_NotifyOutputFileCreated(ByVal pSource As Object, ByVal pFilename As String)
mOutputTickfileName = pFilename
End Sub

Private Sub ITickfileOutputMonitor_NotifyReady(ByVal pSource As Object)
mNotifyReadyCount = mNotifyReadyCount + 1
End Sub

Private Sub ITaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
mTaskCompleted = True
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Test Methods
'@================================================================================

Public Sub WriteTradeBuildv5File()

deleteTradeBuildv5File

Dim lWriter As ITickfileWriter
Set lWriter = mTickfileStore.CreateTickfileWriter(Me, mContractFuture)

Dim lStartTime As Date
lStartTime = Now
Do While mNotifyReadyCount = 0
    Wait 50
    If (Now - lStartTime) > 5 * OneSecond Then Exit Do
Loop

Assert.IsTrue mNotifyReadyCount = 1, "NotifyReady should be called once"

writeTicksA lWriter
lWriter.Finish

Assert.AreEqualStrings TickfileAPath, mOutputTickfileName, , "Incorrect output filename: " & mOutputTickfileName
Assert.IsTrue mNotifyOutputTickfileClosedCount = 1, "NotifyOutputTickfileClosed should be called once"

verifyTickfilev5A

Assert.AreEqual 5, NumberOfReceivedTicks, "Expected 5 ticks, received " & NumberOfReceivedTicks

End Sub

Public Sub WriteTradeBuildv5FileWithBufferedWriter()

deleteTradeBuildv5File

Dim lBufferedWriter As BufferedTickfileWriter
Set lBufferedWriter = CreateBufferedTickfileWriter(mTickfileStore, Me, mContractFuture)

writeTicksA lBufferedWriter
lBufferedWriter.Finish

Wait 50  ' allow time for the NotifyReady and for the buffered ticks to be written

Assert.IsTrue mNotifyReadyCount = 1, "NotifyReady should be called once"
Assert.IsTrue mNotifyOutputTickfileClosedCount = 1, "NotifyOutputTickfileClosed should be called once"
Assert.AreEqualStrings TickfileAPath, mOutputTickfileName, , "Incorrect output filename: " & mOutputTickfileName

verifyTickfilev5A

Assert.AreEqual 5, NumberOfReceivedTicks, "Expected 5 ticks, received " & NumberOfReceivedTicks

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function CreateTickfileSpecWithTradeBuildv5AFile() As TickfileSpecifier
Dim tickfileSpec As New TickfileSpecifier
tickfileSpec.FileName = TickfileAPath
Set CreateTickfileSpecWithTradeBuildv5AFile = tickfileSpec
End Function

Private Sub deleteTradeBuildv5File()
Dim fso As New FileSystemObject
Dim f As File
On Error Resume Next
Set f = fso.GetFile(TickfileAPath)
If Err.Number = 0 Then f.Delete
End Sub

Private Sub verifyTickfilev5A()
mTickfileSpecs.Add CreateTickfileSpecWithTradeBuildv5AFile

Dim tfmev As New TickfileMgrEV
tfmev.Initialise CreateTickfileReplayController(mTickfileStore, , , 0), mTickfileSpecs

tfmev.ReplayController.StartReplay.AddTaskCompletionListener Me

Do
    Wait 50
Loop Until mTaskCompleted

End Sub

Private Sub writeTicksA(ByVal pWriter As ITickfileWriter)
writeTick pWriter, CDate("21/02/2013 08:14:25"), TickTypeAsk, 6720#, 3
writeTick pWriter, CDate("21/02/2013 08:14:26"), TickTypeBid, 6720.5, 5
writeTick pWriter, CDate("21/02/2013 08:14:27") + 120 * OneMillisec, TickTypeTrade, 6720#, 1
writeTick pWriter, CDate("21/02/2013 08:14:27") + 255 * OneMillisec, TickTypeVolume, , 7625
writeTick pWriter, CDate("21/02/2013 08:14:29"), TickTypeClosePrice, 6708#
End Sub

Private Sub writeTick( _
                ByVal pWriter As ITickfileWriter, _
                ByVal pTimestamp As Date, _
                ByVal pTickType As TickTypes, _
                Optional ByVal pPrice As Double, _
                Optional ByVal pSize As Long, _
                Optional ByVal pSide As DOMSides, _
                Optional ByVal pPosition As Long, _
                Optional ByVal pOperation As DOMOperations, _
                Optional ByVal pMarketMaker As String)
pWriter.writeTick gCreateTick(pTimestamp, pTickType, pPrice, pSize, pSide, pPosition, pOperation, pMarketMaker)
End Sub


