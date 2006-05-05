VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTRealtimeDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements IRealtimeDataReader

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const InitialTickQueueCapacity As Long = 1000
Private Const MaxTickTableIndex As Long = 9999
Private Const TimeoutSecs As Long = 10

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type FieldDescriptor
    Name        As String
    DataType    As Long
    size        As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As IRealtimeDataInputServiceConsumer
Private mDataConsumer As IStreamingDataConsumer

Private mContract As TradeBuildSP.IContract
Private mL1QuoteRequestId As Long

Private mName As String

Private WithEvents mQTAPI As QTAPI
Attribute mQTAPI.VB_VarHelpID = -1
Private mPort As Long
Private mServer As String
Private mPassword As String

Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mLogLevel As LogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
Debug.Print "QTSP.QTRealtimeDataReader terminated"
End Sub

'================================================================================
' ITickfileReader Interface Members
'================================================================================

Private Sub IRealtimeDataReader_StartData( _
                ByVal Contract As TradeBuildSP.IContract, _
                ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)

Set mContract = Contract
If mQTAPI Is Nothing Then
    Set mQTAPI = gGetQTAPIInstance(mServer, mPort, mPassword, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection)
    mQTAPI.CommonServiceConsumer = mCommonServiceConsumer
    mQTAPI.logLevel = mLogLevel
End If

If mQTAPI.connectionState = ConnConnected Then
    mL1QuoteRequestId = mQTAPI.subscribeToL1Quotes(mContract, Me)
End If

End Sub

Private Sub IRealtimeDataReader_StartMarketDepth( _
    ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)

End Sub

Private Sub IRealtimeDataReader_StopData()
mQTAPI.unsubscribeFromL1Quotes mL1QuoteRequestId
releaseQTAPI
End Sub

Private Sub IRealtimeDataReader_StopMarketDepth()

End Sub

Private Function IRealtimeDataReader_Supports(ByVal capabilities As Long) As Boolean
IRealtimeDataReader_Supports = False
End Function

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' mQTAPI Event Handlers
'================================================================================

Private Sub mQTAPI_Connected()
mL1QuoteRequestId = mQTAPI.subscribeToL1Quotes(mContract, Me)
End Sub

Private Sub mQTAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
If Not retrying Then
    mServiceConsumer.error StandardSPErrorCodes.TFCantConnectDataSource, _
                            Description
End If
End Sub

'================================================================================
' mTimeoutTimer Event Handlers
'================================================================================

Private Sub mTimeoutTimer_TimerExpired()
releaseQTAPI
mServiceConsumer.error StandardSPErrorCodes.TFDataSourceNotResponding, _
                        "QuoteTracker Server is not responding"
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let CommonServiceConsumer(ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property


Friend Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Public Property Let keepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Friend Property Let logLevel(ByVal value As LogLevels)
mLogLevel = value
End Property

Friend Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Let password(ByVal value As String)
mPassword = value
End Property

Friend Property Let port(ByVal value As Long)
mPort = value
End Property

Friend Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Friend Property Let StreamingDataConsumer(ByVal RHS As IStreamingDataConsumer)
Set mDataConsumer = RHS
End Property

Friend Property Let realtimeDataInputServiceConsumer(ByVal RHS As IRealtimeDataInputServiceConsumer)
Set mServiceConsumer = RHS
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub fireTick( _
                ByRef pTick As Tick)
With pTick
    Select Case .TickType
    Case TickTypes.bid
        mDataConsumer.bid .timestamp, .Price, .size
    Case TickTypes.ask
        mDataConsumer.ask .timestamp, .Price, .size
    Case TickTypes.last
        mDataConsumer.Trade .timestamp, .Price, .size
    Case TickTypes.volume
        mDataConsumer.volume .timestamp, .size
    Case TickTypes.prevClose
        mDataConsumer.PreviousClose .timestamp, .Price
    Case TickTypes.high
        mDataConsumer.high .timestamp, .Price
    Case TickTypes.low
        mDataConsumer.low .timestamp, .Price
    Case TickTypes.openInterest
        mDataConsumer.openInterest .timestamp, .size
    End Select
End With
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub logMessage( _
                ByVal message As String, _
                ByVal logLevel As LogLevels)
If mLogLevel < logLevel Or _
    mLogLevel = LogLevelNone Then Exit Sub
mCommonServiceConsumer.NotifyListeners message, _
                                        TradeBuildSP.StandardListenValueTypes.Log, _
                                        Now
End Sub

Private Sub releaseQTAPI()
If Not mQTAPI Is Nothing Then
    gReleaseQTAPIInstance mQTAPI
    Set mQTAPI = Nothing
End If
End Sub



