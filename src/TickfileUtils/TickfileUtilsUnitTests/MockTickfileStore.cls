VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockTickfileStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDeferredAction
Implements ITickfileStore

Private mDataConsumer                           As IGenericTickListener
Private mReplayMonitor                          As ITickfileReplayMonitor
Private mFormatIdentifier                       As String

Private mAsync                                  As Boolean

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Dim ev As NotificationEventData
ev = Data
fireNotify ev
End Sub

Private Function ITickfileStore_CreateTickfileReader( _
                ByVal pDataConsumer As IGenericTickListener, _
                ByVal pReplayMonitor As ITickfileReplayMonitor, _
                ByVal pStreamId As Long) As ITickfileReader
Set mDataConsumer = pDataConsumer
Set mReplayMonitor = pReplayMonitor

Dim lReader As New MockSyncTFReader
lReader.Initialise pDataConsumer, pReplayMonitor, pStreamId, mAsync
Set ITickfileStore_CreateTickfileReader = lReader
End Function

Private Function ITickfileStore_CreateTickfileWriter(ByVal pOutputMonitor As ITickfileOutputMonitor, ByVal pContractFuture As IFuture, Optional ByVal pFormatIdentifier As String = "", Optional ByVal pLocation As String = "") As ITickfileWriter

End Function

Private Sub ITickfileStore_Finish()

End Sub

Private Property Get ITickfileStore_SupportedFormats() As TickfileFormatSpecifier()

End Property

Private Function ITickfileStore_Supports(ByVal pCapabilities As TickfileStoreCapabilities, Optional ByVal pFormatIdentifier As String) As Boolean

End Function

Private Sub fireNotify(ByRef ev As NotificationEventData)
mReplayMonitor.NotifyEvent ev
End Sub

Friend Property Let Async(ByVal Value As Boolean)
mAsync = Value
End Property

Private Sub fireNotifyEvent(ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Dim lEv As NotificationEventData

Set lEv.Source = Me
lEv.EventCode = pErrorCode
lEv.EventMessage = pErrorMessage
If mAsync Then
    DeferAction Me, lEv, 5, ExpiryTimeUnitMilliseconds
Else
    fireNotify lEv
End If
End Sub



