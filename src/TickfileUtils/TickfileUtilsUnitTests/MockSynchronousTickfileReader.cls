VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockSyncTFReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDeferredAction
Implements ITickfileReader

Private mDataConsumer                           As IGenericTickListener
Private mReplayMonitor                          As ITickfileReplayMonitor

Private mTickfileSpec                           As ITickfileSpecifier

Private mTickIndex                              As Long

Private mStreamId                               As Long

Private mAsync                                  As Boolean

Private Enum DeferredActions
    ActionFireNextTick
    ActionPrepareTickfile
End Enum

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Dim lAction As DeferredActions
lAction = CLng(Data)
Select Case lAction
Case ActionFireNextTick
    FireNextTick
Case ActionPrepareTickfile
    PrepareTickfile
    mReplayMonitor.NotifyTickfilePrepared Me, mTickfileSpec
End Select
End Sub

Private Sub ITickfileReader_CloseInputFile()
' no action needed
End Sub

Private Function ITickfileReader_FireNextTick() As Boolean
If mAsync Then
    DeferAction Me, DeferredActions.ActionFireNextTick
Else
    FireNextTick
    ITickfileReader_FireNextTick = True
End If
End Function

Private Property Get ITickfileReader_FirstTickFileTime() As Date
Select Case mTickfileSpec.FileName
Case "Valid A"
    ITickfileReader_FirstTickFileTime = TicksA(0).TimeStamp
Case Else
End Select
End Property

Private Property Get ITickfileReader_InputTickfileSpecifier() As ITickfileSpecifier

End Property

Private Function ITickfileReader_PrepareTickfile(ByVal pTickfileSpec As ITickfileSpecifier) As Boolean
Set mTickfileSpec = pTickfileSpec
If mAsync Then
    DeferAction Me, DeferredActions.ActionPrepareTickfile
    ITickfileReader_PrepareTickfile = False
Else
    ITickfileReader_PrepareTickfile = PrepareTickfile
End If
End Function

Private Property Get ITickfileReader_ReplayProgressPercent() As Long

End Property

Private Property Get ITickfileReader_ReplayProgressTicks() As Long

End Property

Private Function ITickfileReader_Supports(ByVal Capabilities As Long, Optional ByVal FormatIdentifier As String) As Boolean

End Function

Private Property Get ITickfileReader_TickfileSizeBytes() As Long
Select Case mTickfileSpec.FileName
Case "Valid A"
    ITickfileReader_TickfileSizeBytes = 100
Case "Valid B"
    ITickfileReader_TickfileSizeBytes = 1453
End Select
End Property

Friend Sub Initialise( _
                ByVal pDataConsumer As IGenericTickListener, _
                ByVal pReplayMonitor As ITickfileReplayMonitor, _
                ByVal pStreamId As Long, _
                ByVal pAsync As Boolean)
Set mDataConsumer = pDataConsumer
Set mReplayMonitor = pReplayMonitor
mStreamId = pStreamId
mAsync = pAsync
End Sub

Private Sub FireNextTick()
Dim ev As GenericTickEventData

Select Case mTickfileSpec.FileName
Case "Valid A"
    If mTickIndex <= UBound(TicksA) Then
        Set ev.Source = Me
        ev.StreamId = mStreamId
        ev.Tick = TicksA(mTickIndex)
        mDataConsumer.NotifyTick ev
        mTickIndex = mTickIndex + 1
    Else
        mReplayMonitor.NotifyReplayComplete Me, UBound(TicksA) + 1
    End If
Case "Valid B"
    If mTickIndex <= UBound(TicksB) Then
        Set ev.Source = Me
        ev.StreamId = mStreamId
        ev.Tick = TicksB(mTickIndex)
        mDataConsumer.NotifyTick ev
        mTickIndex = mTickIndex + 1
    Else
        mReplayMonitor.NotifyReplayComplete Me, UBound(TicksB) + 1
    End If
Case Else
End Select
End Sub

Private Function getEventData(ByVal pEvenCode As Long, ByVal pEventMessage As String) As NotificationEventData
Set getEventData.Source = Me
getEventData.EventCode = pEvenCode
getEventData.EventMessage = pEventMessage
End Function

Private Function PrepareTickfile() As Boolean
Select Case mTickfileSpec.FileName
Case "Valid A"
    mTickfileSpec.Contract = gCreateContractFromLocalSymbol("ZM13")
    mReplayMonitor.NotifyReady Me, mTickfileSpec
    PrepareTickfile = True
Case "Valid B"
    mTickfileSpec.Contract = gCreateContractFromLocalSymbol("ESM3")
    mReplayMonitor.NotifyReady Me, mTickfileSpec
    PrepareTickfile = True
Case "Empty file"
    mReplayMonitor.NotifyEvent getEventData(TickfileEventCodes.TickfileEventNoTickData, "Specified tickfile is empty")
    PrepareTickfile = True
Case "Invalid file"
    mReplayMonitor.NotifyEvent getEventData(TickfileEventCodes.TickfileEventFileIsInvalid, "Specified tickfile is invalid")
    mReplayMonitor.NotifyReplayComplete Me, 0
    PrepareTickfile = False
Case Else
    mReplayMonitor.NotifyEvent getEventData(TickfileEventCodes.TickfileEventFileDoesNotExist, "Specified tickfile doesn't exist")
    PrepareTickfile = False
End Select
End Function


