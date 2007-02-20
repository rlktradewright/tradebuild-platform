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


'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mRealtimeDataServiceProvider As QTRealtimeDataServiceProvider

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

On Error GoTo err
Set mContract = Contract
If mQTAPI Is Nothing Then
    Set mQTAPI = gGetQTAPIInstance(mServer, mPort, mPassword, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection)
    mQTAPI.CommonServiceConsumer = mCommonServiceConsumer
    mQTAPI.logLevel = mLogLevel
End If

If mQTAPI.connectionState = ConnConnected Then
    mL1QuoteRequestId = mQTAPI.subscribeToL1Quotes(mContract, Me)
End If
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTRealtimeDataReader" & "::" & "IRealtimeDataReader_StartData", _
                err.Description

End Sub

Private Sub IRealtimeDataReader_StartMarketDepth( _
    ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)

End Sub

Private Sub IRealtimeDataReader_StopData()
On Error GoTo err
finish
mRealtimeDataServiceProvider.dataReaderFinished Me
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTRealtimeDataReader" & "::" & "IRealtimeDataReader_StopData", _
                err.Description
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
On Error GoTo err
mServiceConsumer.NotifyEvent StandardSPEventCodes.RTConnectedToDataSource, _
                        "Connected to data source", _
                        mRealtimeDataServiceProvider.handle

mL1QuoteRequestId = mQTAPI.subscribeToL1Quotes(mContract, Me)
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTRealtimeDataReader" & "::" & "mQTAPI_Connected", _
                err.Description
End Sub

Private Sub mQTAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
On Error GoTo err
If retrying Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.RTRetryConnectDataSource, _
                            Description, _
                            mRealtimeDataServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.RTCantConnectDataSource, _
                            Description, _
                            mRealtimeDataServiceProvider.handle
End If
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTRealtimeDataReader" & "::" & "mQTAPI_ConnectFailed", _
                err.Description
End Sub

Private Sub mQTAPI_ConnectionClosed(ByVal reconnecting As Boolean)
If reconnecting Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.RTReConnectingDataSource, _
                        "Lost connection to QT (reconnecting)", _
                        mRealtimeDataServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.RTLostConnectionToDataSource, _
                        "Lost connection to QT", _
                        mRealtimeDataServiceProvider.handle
End If
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

Friend Property Let realtimeDataServiceProvider( _
                ByVal value As QTRealtimeDataServiceProvider)
Set mRealtimeDataServiceProvider = value
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

Friend Sub finish()
If Not mQTAPI Is Nothing Then
    mQTAPI.unsubscribeFromL1Quotes mL1QuoteRequestId
    gReleaseQTAPIInstance mQTAPI
    Set mQTAPI = Nothing
End If
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
Set mContract = Nothing
End Sub

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

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError Number, _
                Source, _
                Description, _
                mRealtimeDataServiceProvider.handle

finish
mRealtimeDataServiceProvider.dataReaderFinished Me
End Sub

Private Sub logMessage( _
                ByVal message As String, _
                ByVal logLevel As LogLevels)
If mLogLevel < logLevel Or _
    mLogLevel = LogLevelNone Then Exit Sub
mCommonServiceConsumer.NotifyListeners message, _
                                        TradeBuildSP.StandardListenValueTypes.Log, _
                                        Now
End Sub



